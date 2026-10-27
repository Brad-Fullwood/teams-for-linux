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
 * rewrites the encodings Teams capped at 720p to the configured maximum size
 * and bitrate.
 *
 * In meetings Teams sends simulcast: rid 1 is a 320x180 thumbnail at
 * 133 kbps that every receiver gets, rid 2 is the main layer that Teams sets
 * to 720p, 540p or 180p from what the receivers show, rid 3 is unused. Only
 * a layer Teams itself set at its 720p cap is lifted. The lower layers keep
 * Teams' values. Lifting the thumbnail as well (every active layer, the rule
 * until 2026-10-27) put two 1080p layers under the service's 3.9 Mbps cap:
 * libwebrtc fed the thumbnail first, the main layer encoded nothing, and the
 * other participants saw a white tile.
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

// A layer Teams capped at this height or above is the one it meant as the
// full picture; anything smaller is a deliberate downscale for small tiles.
const LIFT_FROM_HEIGHT = 720;

const INACTIVE_DOWN_TO = { maxWidth: 640, maxHeight: 360 };

const applyCameraSendResolutionPatch = function (senderConfig) {
  const maxWidth = senderConfig.maxWidth || DEFAULT_MAX_WIDTH;
  const maxHeight = senderConfig.maxHeight || DEFAULT_MAX_HEIGHT;
  const maxBitrate = senderConfig.maxBitrate || DEFAULT_MAX_BITRATE;
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

  function shouldLift(enc) {
    return enc.active !== false && layerHeight(enc) >= LIFT_FROM_HEIGHT;
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
      const lifting = encodings.some(shouldLift);
      const anyDownTo = encodings.some((enc) => enc.scaleResolutionDownTo);
      const rewritten = encodings.map((enc) => {
        const copy = { ...enc };
        if (shouldLift(enc)) {
          copy.scaleResolutionDownTo = { maxWidth, maxHeight };
          copy.scaleResolutionDownBy = 1;
          copy.maxBitrate = maxBitrate;
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
