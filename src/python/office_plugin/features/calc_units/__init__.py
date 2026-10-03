"""Units to Numbers: 1 lakh, 1L, 2 cr, 2cr, 1.5k, 3M ... -> numbers in the selected cells."""

from office_plugin.core.feature import Feature
from office_plugin.services import range_convert as RC
from office_plugin.services.numunits import parse_units


class CalcUnitsFeature(Feature):
    def execute(self, command, frame, args):
        docs = self.services.get("doc_utils")
        ui = self.services.get("ui")
        controller = docs.controller_of(frame) or docs.active_controller()
        if controller is None:
            return
        title = "Units to Numbers"
        answer = ui.form(title, [
            {"type": "label", "text": "Convert amounts like 1 lakh, 1L, 2 cr, 2cr, 1.5k, 3M, Rs. 2 crore,\n"
                                      "12,34,567 or 1 crore 25 lakh into numbers. Numbers are left alone."},
            {"type": "dropdown", "name": "delivery", "label": "Put the numbers",
             "items": list(RC.DELIVERIES), "value": RC.RIGHT},
        ], frame)
        if answer is None:
            return
        RC.run(ui, frame, controller, title,
               accepts=lambda data, formula: isinstance(data, str) and data.strip() != "",
               convert=lambda text: float(parse_units(text)), delivery=answer["delivery"],
               nothing_text="The selection contains no text to read.", done_noun="amount(s) converted to numbers")
