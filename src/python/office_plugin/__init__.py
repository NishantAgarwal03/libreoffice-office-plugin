"""Office Plugin for LibreOffice: a modular bundle of MS-Office-style features.

Layout:
    core/      infrastructure only (events, registry, dispatch routing, settings)
    services/  shared helpers features may reuse (positions, ui, doc_utils)
    features/  one folder per user-facing feature; each has a feature.json

Modules here must stay importable WITHOUT the `uno` module (so the pure logic
can be unit-tested with any Python); anything that needs `uno` imports it lazily.
"""

__version__ = "0.7.2"
