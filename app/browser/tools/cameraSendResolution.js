/**
 * CameraSendResolution Browser Tool
 *
 * Raises the resolution and bitrate cap that Microsoft Teams puts on the
 * outgoing camera video. Teams calls RTCRtpSender.setParameters with
 * encodings capped at 1280x720 (scaleResolutionDownTo, scaleResolutionDownBy
 * and maxBitrate), so the camera is sent at 720p or less even when it
 * captures at 1080p.
 *
 * This tool intercepts setParameters and addTransceiver for camera tracks and
 * rewrites the main layer to the configured maximum size, bitrate and frame
 * rate.
 *
 * Teams sends three simulcast encodings. With two people in a meeting only
 * rid 1 is active and it is the main layer, capped at 720p. With more people
 * rid 1 is a 320x180 thumbnail at 133 kbps that every receiver gets and rid 2
 * is the main layer, which Teams sets to 720p, 540p or 180p from what the
 * receivers show. rid 3 is unused. The main layer is the active encoding with
 * the largest cap, the later one on a tie. The first encoding at thumbnail
 * size (180p or less) is never lifted, so a meeting where only the thumbnail
 * is active sends Teams' values. Lifting the thumbnail as well (every active
 * layer, the rule until 2026-10-27) put two 1080p layers under the service's
 * 3.9 Mbps cap: libwebrtc fed the thumbnail first, the main layer encoded
 * nothing, and the other participants saw a white tile. With one encoding
 * (1:1 calls) that encoding is the main layer.
 *
 * liftFromHeight (default 720) is the smallest cap Teams must have set on the
 * main layer for it to be lifted. 0 lifts it whatever size Teams asked for,
 * so receivers showing a small tile get 1080p as well.
 *
 * An encoding without an `active` key is active (the WebRTC default).
 * Chromium rejects parameters where scaleResolutionDownTo is set on some
 * encodings but not others, so when the lifted layer gains one and the
 * others had none, they get one that matches their scaleResolutionDownBy.
 *
 * The remote SDP from the Teams service carries b=AS:3920 on every video
 * section, which caps Chromium's outgoing bandwidth estimate at 3.92 Mbps.
 * When remoteBandwidthKbps is above 0, setRemoteDescription rewrites every
 * b=AS line to that value and every b=TIAS line to the same value in bps.
 */

const DEFAULT_MAX_WIDTH = 1920;
const DEFAULT_MAX_HEIGHT = 1080;
const DEFAULT_MAX_BITRATE = 8000000;

const DEFAULT_MAX_FRAMERATE = 30;
const DEFAULT_LIFT_FROM_HEIGHT = 720;
const THUMBNAIL_HEIGHT = 180;

const INACTIVE_DOWN_TO = { maxWidth: 640, maxHeight: 360 };

const applyCameraSendResolutionPatch = function (senderConfig) {
  const maxWidth = senderConfig.maxWidth || DEFAULT_MAX_WIDTH;
  const maxHeight = senderConfig.maxHeight || DEFAULT_MAX_HEIGHT;
  const maxBitrate = senderConfig.maxBitrate || DEFAULT_MAX_BITRATE;
  const maxFramerate = senderConfig.maxFramerate || DEFAULT_MAX_FRAMERATE;
  const liftFromHeight =
    senderConfig.liftFromHeight ?? DEFAULT_LIFT_FROM_HEIGHT;
  const remoteBandwidthKbps = senderConfig.remoteBandwidthKbps || 0;

  // Screen-share tracks expose displaySurface in their settings
  function isCameraTrack(track) {
    return track?.kind === "video" && !track.getSettings?.().displaySurface;
  }

  // The height Teams capped a layer at: scaleResolutionDownTo when set,
  // else the capture height over scaleResolutionDownBy, else unlimited.
  function layerHeight(enc) {
    const to = enc.scaleResolutionDownTo;
    if (to && (to.maxHeight || to.height)) {
      return to.maxHeight || to.height;
    }
    if (enc.scaleResolutionDownBy > 1) {
      return maxHeight / enc.scaleResolutionDownBy;
    }
    return Infinity;
  }

  // Index of the main layer to lift, or -1: the active encoding with the
  // largest cap, the later one on a tie. The first of several encodings at
  // thumbnail size is the meeting thumbnail and is not lifted.
  function mainLayerIndex(encodings) {
    let best = -1;
    for (let i = 0; i < encodings.length; i++) {
      const enc = encodings[i];
      if (enc.active === false) {
        continue;
      }
      if (best === -1 || layerHeight(enc) >= layerHeight(encodings[best])) {
        best = i;
      }
    }
    if (best === -1 || layerHeight(encodings[best]) < liftFromHeight) {
      return -1;
    }
    const isThumbnail =
      best === 0 &&
      encodings.length > 1 &&
      layerHeight(encodings[0]) <= THUMBNAIL_HEIGHT;
    return isThumbnail ? -1 : best;
  }

  // scaleResolutionDownTo for a layer that is not lifted, when the lifted
  // layer needs one on every encoding.
  function downToFor(enc) {
    const by = enc.scaleResolutionDownBy;
    if (by > 1) {
      return {
        maxWidth: Math.round(maxWidth / by),
        maxHeight: Math.round(maxHeight / by),
      };
    }
    return { ...INACTIVE_DOWN_TO };
  }

  // Rewrites the encodings array in place. Returns false when it fails so the
  // caller can pass the original parameters through.
  function rewriteEncodings(encodings, source) {
    if (!Array.isArray(encodings) || encodings.length === 0) {
      return true;
    }

    try {
      const original = JSON.stringify(encodings);
      const main = mainLayerIndex(encodings);
      const lifting = main !== -1;
      const anyDownTo = encodings.some((enc) => enc.scaleResolutionDownTo);
      const rewritten = encodings.map((enc, i) => {
        const copy = { ...enc };
        if (i === main) {
          copy.scaleResolutionDownTo = { maxWidth, maxHeight };
          copy.scaleResolutionDownBy = 1;
          copy.maxBitrate = maxBitrate;
          copy.maxFramerate = Math.max(enc.maxFramerate || 0, maxFramerate);
        } else if (lifting && !anyDownTo) {
          copy.scaleResolutionDownTo = downToFor(enc);
        }
        return copy;
      });

      console.debug(
        `[CAMERA_SEND_RESOLUTION] ${source} original encodings:`,
        original,
      );
      encodings.splice(0, encodings.length, ...rewritten);
      console.debug(
        `[CAMERA_SEND_RESOLUTION] ${source} modified encodings:`,
        JSON.stringify(encodings),
      );
      return true;
    } catch (error) {
      console.error(
        `[CAMERA_SEND_RESOLUTION] Failed to rewrite ${source} encodings, passing original through:`,
        error,
      );
      return false;
    }
  }

  const originalSetParameters = RTCRtpSender.prototype.setParameters;
  RTCRtpSender.prototype.setParameters = function setParameters(
    parameters,
    ...rest
  ) {
    if (isCameraTrack(this.track) && parameters?.encodings) {
      // Work on a copy so a failed rewrite leaves the caller's object intact
      const copy = { ...parameters, encodings: [...parameters.encodings] };
      if (rewriteEncodings(copy.encodings, "setParameters")) {
        parameters = copy;
      }
    }
    return originalSetParameters.call(this, parameters, ...rest);
  };
  console.debug("[CAMERA_SEND_RESOLUTION] Patched RTCRtpSender.setParameters");

  const originalAddTransceiver = RTCPeerConnection.prototype.addTransceiver;
  RTCPeerConnection.prototype.addTransceiver = function addTransceiver(
    trackOrKind,
    init,
    ...rest
  ) {
    if (
      typeof trackOrKind === "object" &&
      isCameraTrack(trackOrKind) &&
      init?.sendEncodings
    ) {
      const copy = { ...init, sendEncodings: [...init.sendEncodings] };
      if (rewriteEncodings(copy.sendEncodings, "addTransceiver")) {
        init = copy;
      }
    }
    return originalAddTransceiver.call(this, trackOrKind, init, ...rest);
  };
  console.debug(
    "[CAMERA_SEND_RESOLUTION] Patched RTCPeerConnection.addTransceiver",
  );

  if (remoteBandwidthKbps > 0) {
    patchRemoteBandwidth(remoteBandwidthKbps);
  }

  console.debug(
    `[CAMERA_SEND_RESOLUTION] Successfully initialized - cap: ${maxWidth}x${maxHeight} at ${maxBitrate} bps`,
  );
};

// Rewrites b=AS and b=TIAS in the remote SDP. A description without an sdp
// string (implicit rollback) and any failure pass through unchanged.
function patchRemoteBandwidth(kbps) {
  const originalSetRemoteDescription =
    RTCPeerConnection.prototype.setRemoteDescription;
  RTCPeerConnection.prototype.setRemoteDescription =
    function setRemoteDescription(description, ...rest) {
      try {
        const sdp = description?.sdp;
        if (typeof sdp === "string" && /b=(AS|TIAS):\d+/.test(sdp)) {
          const replaced = [];
          const newSdp = sdp
            .replaceAll(/b=AS:(\d+)/g, (match) => {
              replaced.push(match);
              return `b=AS:${kbps}`;
            })
            .replaceAll(/b=TIAS:(\d+)/g, (match) => {
              replaced.push(match);
              return `b=TIAS:${kbps * 1000}`;
            });
          console.debug(
            `[CAMERA_SEND_RESOLUTION] Remote SDP bandwidth ${[...new Set(replaced)].join(", ")} rewritten to b=AS:${kbps}`,
          );
          description = { type: description.type, sdp: newSdp };
        }
      } catch (error) {
        console.error(
          "[CAMERA_SEND_RESOLUTION] Failed to rewrite remote SDP bandwidth, passing original through:",
          error,
        );
      }
      return originalSetRemoteDescription.call(this, description, ...rest);
    };
  console.debug(
    "[CAMERA_SEND_RESOLUTION] Patched RTCPeerConnection.setRemoteDescription",
  );
}

/**
 * Initialize the cameraSendResolution tool
 * @param {Object} config - Application configuration
 */
function init(config) {
  const senderConfig = config.media?.camera?.sender;
  if (!senderConfig?.enabled) {
    return;
  }

  try {
    applyCameraSendResolutionPatch(senderConfig);
    console.info(
      `[CAMERA_SEND_RESOLUTION] Outgoing camera video cap raised to ${senderConfig.maxWidth || DEFAULT_MAX_WIDTH}x${senderConfig.maxHeight || DEFAULT_MAX_HEIGHT}`,
    );
  } catch (error) {
    console.error("[CAMERA_SEND_RESOLUTION] Failed to initialize:", error);
  }
}

module.exports = { init };
