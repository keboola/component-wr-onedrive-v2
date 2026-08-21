"""Unit coverage for the VCR recording-support mechanisms (``src/vcr_sanitizers.py``).

This file mirrors ``component-ex-medallia``'s ``tests/test_recording_support.py`` slot in the
canonical layout: dedicated, network-free coverage of the sanitizer classes that shape every
committed cassette under ``tests/functional/`` — ``DefaultSanitizer``'s patched sensitive-field
list, ``QueryParamSanitizer(parameters=["tempauth"])``, ``_IdentityFieldRedactor``,
``_GuidRedactor`` (GUID collapsing to a fixed placeholder), and ``_StreamedBodySerializationFix``
(the streamed file-body cassette-serialization workaround) — see that module's own docstring and
class-level comments for the full mechanics.

As of this reorganization, ``src/vcr_sanitizers.py`` has no dedicated unit tests of its own in
this repository: it is currently exercised only indirectly, through ``tests/test_functional.py``
replaying the already-sanitized cassettes under ``tests/functional/`` (and, at record time, by
``tests/setup/record_vcr_cassettes.py``'s cassette-validation-gate step — see
``tests/functional/README.md``). This is a deliberate gap noted during the reorganization, not
something papered over here with fabricated tests (the restructure that introduced this file is a
pure reorganization — no test may be invented, deleted, or weakened). Direct unit tests for each
sanitizer class (mirroring ``component-ex-medallia``'s allowlist-style
``MedalliaResponseBodySanitizer`` coverage) belong in this module once written.
"""
