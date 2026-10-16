# `fixtures/v1_parity/`

Verbatim expected outputs from `keboola.wr-onedrive` (v1, PHP)'s own datadir test suite
(`tests/datadir/<case>/` in that repo) — `config.json` plus `expected-stdout` or
`expected-code`/`expected-stderr`, copied byte-for-byte, one directory per case.

`tests/test_v1_parity.py` builds a `Component` with a mocked `GraphClient`, invokes the relevant
sync action directly, and asserts the result matches the golden fixture here (modulo v1's own
`%s`/`%a`/`%A` wildcards for values that are dynamic on v1's real test tenant — drive ids, file
ids, worksheet ids — see that module's docstring for the wildcard-matching details).

Matching these byte-for-byte is an acceptance criterion for the four sync actions ported from v1
— `search`, `createWorkbook`, `createWorksheet`, `getWorksheets` must keep producing the exact
output shape v1 did, so that existing Keboola configurations relying on that output don't break
when a project migrates from v1 to v2.

Do not hand-edit these fixtures. If v1's own fixtures change, re-copy them from
`keboola.wr-onedrive`'s test suite verbatim.
