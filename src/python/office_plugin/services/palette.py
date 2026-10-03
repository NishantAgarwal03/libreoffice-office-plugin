"""Cell background colours the plugin paints (pure data).

Purpose: one palette shared by every tool that colours cells (Highlight, placeholder rows...), so
         "Clear Highlight" can recognise and remove exactly the plugin's own colours.
Inputs/Outputs: name -> 0xRRGGBB integer.
Must not: change values between versions (users' documents already contain them).
"""

PALETTE = {"Yellow": 0xFFEB84, "Green": 0xA9E5A0, "Orange": 0xFFC285,
           "Red": 0xFF9E9E, "Blue": 0x9DC3FF, "Purple": 0xD1B3FF}
