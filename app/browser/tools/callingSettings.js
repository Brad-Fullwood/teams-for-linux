/**
 * CallingSettings Browser Tool
 *
 * Merges settings over the configuration of the Teams web calling stack, the
 * way Teams' own devtools API does (`webMA.config.setOverrides`, which Teams
 * attaches only when its remote config sets `enableDevtoolsAPI`).
 *
 * The calling stack keeps its settings in a config provider object whose
 * constructor runs `this._consoleOverrides = {}` and which merges that object
 * over every settings update (`mergeSettingsWithConsoleOverrides`). This tool
 * puts a setter for `_consoleOverrides` on Object.prototype before Teams loads.
 * The constructor's assignment reaches the setter, which stores the configured
 * overrides on the instance as a plain property, removes itself, and then asks
 * Teams to recompute its settings through `setConsoleOverrides`.
 *
 * A call takes a copy of the settings when its pre-join screen opens, so the
 * overrides apply to every call started after Teams has loaded.
 *
 * Main use: the outgoing camera video. With the defaults Teams' web client
 * tells the service it sends at most 1280x720 (frame size 3600 macroblocks in
 * webrtcVideoCapabilityMaxFS, webrtcCameraOpenFs, hardwareEncodedVideoMaxFS
 * and sendVideoCapabilityManager.defaultFS). Raising those to 8160 and setting
 * video1080p.software.isEnabled makes Teams offer and send 1920x1080, and the
 * service forwards it to receivers that show a large tile.
 */

const PROPERTY = "_consoleOverrides";

function init(config) {
  const overrides = config.media?.callingSettings;
  if (
    !overrides ||
    typeof overrides !== "object" ||
    Object.keys(overrides).length === 0
  ) {
    return;
  }

  if (Object.hasOwn(Object.prototype, PROPERTY)) {
    console.warn("[CALLING_SETTINGS] Hook already installed, skipping");
    return;
  }

  Object.defineProperty(Object.prototype, PROPERTY, {
    configurable: true,
    enumerable: false,
    get() {
      return undefined;
    },
    set(value) {
      // The first assignment comes from the config provider's constructor.
      // Later assignments go to the instance's own property defined below.
      delete Object.prototype[PROPERTY];
      Object.defineProperty(this, PROPERTY, {
        configurable: true,
        enumerable: true,
        writable: true,
        value: { ...value, ...structuredClone(overrides) },
      });
      const provider = this;
      // Runs after the constructor has finished
      setTimeout(() => {
        try {
          if (typeof provider.setConsoleOverrides === "function") {
            provider.setConsoleOverrides({});
            console.info(
              `[CALLING_SETTINGS] Applied ${Object.keys(overrides).length} calling setting override(s)`,
            );
          } else {
            console.warn(
              "[CALLING_SETTINGS] Captured object has no setConsoleOverrides, overrides apply on the next settings update only",
            );
          }
        } catch (error) {
          console.error("[CALLING_SETTINGS] Failed to apply overrides:", error);
        }
      }, 0);
    },
  });
  console.debug("[CALLING_SETTINGS] Waiting for the calling config provider");
}

module.exports = { init };
