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
 * rewrites every active encoding to the configured maximum size and bitrate.
 * Inactive encodings keep their values. Chromium rejects parameters where
 * scaleResolutionDownTo is set on some encodings but not others, so when no
 * encoding carries it the inactive ones get a 640x360 value as well.
 *
 * The remote SDP from the Teams service carries b=AS:3920 on every video
 * section, which caps Chromium's outgoing bandwidth estimate at 3.92 Mbps.
 * When remoteBandwidthKbps is above 0, setRemoteDescription rewrites every
 * b=AS line to that value and every b=TIAS line to the same value in bps.
 */

const DEFAULT_MAX_WIDTH = 1920;
const DEFAULT_MAX_HEIGHT = 1080;
const DEFAULT_MAX_BITRATE = 8000000;

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

  // Rewrites the encodings array in place. Returns false when it fails so the
  // caller can pass the original parameters through.
  function rewriteEncodings(encodings, source) {
    if (!Array.isArray(encodings) || encodings.length === 0) {
      return true;
    }

    try {
      const original = JSON.stringify(encodings);
      const anyDownTo = encodings.some((enc) => enc.scaleResolutionDownTo);
      const rewritten = encodings.map((enc) => {
        const copy = { ...enc };
        if (copy.active === true) {
          copy.scaleResolutionDownTo = { maxWidth, maxHeight };
          copy.scaleResolutionDownBy = 1;
          copy.maxBitrate = maxBitrate;
        } else if (!anyDownTo) {
          copy.scaleResolutionDownTo = { ...INACTIVE_DOWN_TO };
        }
        return copy;
      });

      console.debug(
        `[CAMERA_SEND_RESOLUTION] ${source} original encodings:`,
        original
      );
      encodings.splice(0, encodings.length, ...rewritten);
      console.debug(
        `[CAMERA_SEND_RESOLUTION] ${source} modified encodings:`,
        JSON.stringify(encodings)
      );
      return true;
    } catch (error) {
      console.error(
        `[CAMERA_SEND_RESOLUTION] Failed to rewrite ${source} encodings, passing original through:`,
        error
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
    "[CAMERA_SEND_RESOLUTION] Patched RTCPeerConnection.addTransceiver"
  );

  if (remoteBandwidthKbps > 0) {
    patchRemoteBandwidth(remoteBandwidthKbps);
  }

  console.debug(
    `[CAMERA_SEND_RESOLUTION] Successfully initialized - cap: ${maxWidth}x${maxHeight} at ${maxBitrate} bps`
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
            `[CAMERA_SEND_RESOLUTION] Remote SDP bandwidth ${[...new Set(replaced)].join(", ")} rewritten to b=AS:${kbps}`
          );
          description = { type: description.type, sdp: newSdp };
        }
      } catch (error) {
        console.error(
          "[CAMERA_SEND_RESOLUTION] Failed to rewrite remote SDP bandwidth, passing original through:",
          error
        );
      }
      return originalSetRemoteDescription.call(this, description, ...rest);
    };
  console.debug(
    "[CAMERA_SEND_RESOLUTION] Patched RTCPeerConnection.setRemoteDescription"
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
      `[CAMERA_SEND_RESOLUTION] Outgoing camera video cap raised to ${senderConfig.maxWidth || DEFAULT_MAX_WIDTH}x${senderConfig.maxHeight || DEFAULT_MAX_HEIGHT}`
    );
  } catch (error) {
    console.error("[CAMERA_SEND_RESOLUTION] Failed to initialize:", error);
  }
}

module.exports = { init };
