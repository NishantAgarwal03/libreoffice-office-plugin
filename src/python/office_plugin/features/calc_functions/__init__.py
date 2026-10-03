"""Number Functions: opt-in live cell functions NUMTOWORDS / WORDSTONUM / UNITSTONUM.

Enable writes the generated Basic module (basic_source.py) into the user's My Macros > Standard
library; the Basic wrappers call the plugin's FunctionJob component, which runs the Python logic.
While enabled the module is re-synced at every start so it follows plugin updates. Disable removes it.
Remove the module (Disable) before uninstalling the extension: leftover functions would only return #VALUE!.
"""

from office_plugin.core.feature import Feature

from . import basic_source as B

_TITLE = "Number Functions"
_ENABLE, _DISABLE = "Enable the functions", "Disable (remove them)"
_LIBRARY = "Standard"


class CalcFunctionsFeature(Feature):
    def start(self):
        if self.option("installed", False):               # keep the installed module in step with this plugin version
            try:
                self.install()
            except Exception:
                self.log.exception("calc_functions: could not refresh the Basic module")

    # ---- the module in My Macros ---------------------------------------------------------------
    def _library(self):
        libs = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.script.ApplicationScriptLibraryContainer", self.ctx)
        if not libs.hasByName(_LIBRARY):
            libs.createLibrary(_LIBRARY)
        libs.loadLibrary(_LIBRARY)
        return libs, libs.getByName(_LIBRARY)

    def is_installed(self):
        _, lib = self._library()
        return lib.hasByName(B.MODULE_NAME)

    def install(self):
        """Create or update the module; returns 'installed' | 'updated' | 'unchanged'."""
        libs, lib = self._library()
        text = B.source()
        if lib.hasByName(B.MODULE_NAME):
            if lib.getByName(B.MODULE_NAME) == text:
                return "unchanged"
            lib.replaceByName(B.MODULE_NAME, text)
            result = "updated"
        else:
            lib.insertByName(B.MODULE_NAME, text)
            result = "installed"
        self._store(libs)
        return result

    def uninstall(self):
        libs, lib = self._library()
        if not lib.hasByName(B.MODULE_NAME):
            return False
        lib.removeByName(B.MODULE_NAME)
        self._store(libs)
        return True

    @staticmethod
    def _store(libs):
        try:
            libs.storeLibraries()                          # write to the profile now, not only at exit
        except Exception:
            pass

    # ---- command -----------------------------------------------------------------------------------
    def execute(self, command, frame, args):
        ui = self.services.get("ui")
        state = "enabled" if self._safe_installed() else "not enabled"
        answer = ui.form(_TITLE, [
            {"type": "label", "text": "Live cell functions (currently %s):\n"
                                      "  =NUMTOWORDS(number; [system]; [currency]; [letters])\n"
                                      "      system: Indian (default) / International; currency: INR, USD, EUR, GBP;\n"
                                      "      letters: sentence, title, upper, lower\n"
                                      "  =WORDSTONUM(text)      =UNITSTONUM(text)   e.g. 2cr, 1 lakh, 1.5k\n\n"
                                      "Enabling adds a small module to My Macros > Standard. Disable it\n"
                                      "before uninstalling the extension." % state},
            {"type": "dropdown", "name": "action", "label": "Action", "items": [_ENABLE, _DISABLE],
             "value": _DISABLE if state == "enabled" else _ENABLE},
        ], frame)
        if answer is None:
            return
        try:
            if answer["action"] == _ENABLE:
                result = self.install()
                self.settings.set_option(self.id, "installed", True)
                self.settings.save()
                ui.message(_TITLE, "The functions are %s. Use =NUMTOWORDS(A1), =WORDSTONUM(A1), =UNITSTONUM(A1)." %
                           {"installed": "now available", "updated": "updated and available",
                            "unchanged": "already available"}[result], frame)
            else:
                removed = self.uninstall()
                self.settings.set_option(self.id, "installed", False)
                self.settings.save()
                ui.message(_TITLE, "The functions were removed." if removed else "The functions were not installed.", frame)
        except Exception as exc:
            self.log.exception("calc_functions failed")
            ui.message(_TITLE, "Could not change My Macros: %s" % exc, frame, "error")

    def _safe_installed(self):
        try:
            return self.is_installed()
        except Exception:
            return False
