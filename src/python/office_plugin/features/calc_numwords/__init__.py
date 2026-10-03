"""Number to Words / Words to Number for the selected cells (Indian or international wording)."""

from office_plugin.core.feature import Feature
from office_plugin.services import numwords as N
from office_plugin.services import range_convert as RC

_SYSTEMS = {"Indian (lakh, crore)": N.INDIAN, "International (million, billion)": N.INTERNATIONAL}
_PLAIN = "Plain number"
_FORMATS = [_PLAIN] + ["Currency: " + name for name in N.CURRENCIES]
_CASES = [N.SENTENCE, N.TITLE, N.UPPER, N.LOWER]


class CalcNumWordsFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        if command == "ToWords":
            return self._to_words(ui, frame, controller)
        return self._from_words(ui, frame, controller)

    def _to_words(self, ui, frame, controller):
        title = "Number to Words"
        answer = ui.form(title, [
            {"type": "label", "text": "Write the selected numbers in words."},
            {"type": "dropdown", "name": "system", "label": "Numbering system",
             "items": list(_SYSTEMS), "value": "Indian (lakh, crore)"},
            {"type": "dropdown", "name": "format", "label": "Format", "items": _FORMATS, "value": _PLAIN},
            {"type": "dropdown", "name": "case", "label": "Letters", "items": _CASES, "value": N.SENTENCE},
            {"type": "checkbox", "name": "use_and", "value": False,
             "label": "Use \"and\" (one hundred and five)"},
            {"type": "checkbox", "name": "only", "value": True,
             "label": "Add \"Only\" after currency amounts"},
            {"type": "dropdown", "name": "delivery", "label": "Put the words",
             "items": list(RC.DELIVERIES), "value": RC.RIGHT},
        ], frame)
        if answer is None:
            return
        system = _SYSTEMS[answer["system"]]
        currency = None if answer["format"] == _PLAIN else answer["format"].split(": ", 1)[1]

        def convert(value):
            return N.to_words(value, system, currency, answer["case"], answer["use_and"], answer["only"])

        RC.run(ui, frame, controller, title,
               accepts=lambda data, formula: isinstance(data, float),
               convert=convert, delivery=answer["delivery"],
               nothing_text="The selection contains no numbers.", done_noun="number(s) written in words",
               skip_special_formats=True)       # dates, times and TRUE/FALSE are not amounts

    def _from_words(self, ui, frame, controller):
        title = "Words to Number"
        answer = ui.form(title, [
            {"type": "label", "text": "Read number words in the selected cells (Indian or international;\n"
                                      "e.g. \"Rupees twelve lakh and Paise fifty Only\")."},
            {"type": "dropdown", "name": "delivery", "label": "Put the numbers",
             "items": list(RC.DELIVERIES), "value": RC.RIGHT},
        ], frame)
        if answer is None:
            return
        RC.run(ui, frame, controller, title,
               accepts=lambda data, formula: isinstance(data, str) and data.strip() != "",
               convert=lambda text: float(N.parse_words(text)), delivery=answer["delivery"],
               nothing_text="The selection contains no text to read.", done_noun="text(s) converted to numbers")
