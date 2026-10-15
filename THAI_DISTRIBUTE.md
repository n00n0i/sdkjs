# Thai Distributed alignment (Document Server 9.4.0-129)

Word paragraph alignment `w:jc w:val="thaiDistribute"` spreads leftover line width across spaces and Thai grapheme clusters, and it breaks Thai lines on words (`ดูการบ้าน`, not `ดูกา|รบ้าน`). Document Server 9.4.0-129 does neither until both **x2t** (ONLYOFFICE/core) and **sdkjs** are replaced. The screenshots of the Thai contract (สัญญาซื้อขาย, TH SarabunPSK 16) show the old result: a few phrase spaces take the entire extra width, and the line breaks do not match Word.

Package version `9.4.0-129` is git tag `v9.4.0.129` in both repositories.

| Repo | Tag commit | Where the change is |
| --- | --- | --- |
| [n00n0i/core](https://github.com/n00n0i/core) | `a016fc2868b9f99f59dd30aa2088e68bbe5104af` | `core-patches/thai-distribute.patch` (push was denied; see below) |
| [n00n0i/sdkjs](https://github.com/n00n0i/sdkjs) | `75b53d504bfd81bfa35d59bb57dd9e56f7c9235d` | `cursor/thai-distribute-9.4.0-703e` (deploy this one) |

The sdkjs pull request into `master` is branch `cursor/thai-distribute-master-703e`. Word editor sources on that tag and on `master` match; `master` only deletes `tests/` and exports `ApiWorkbook`. The pull request URL is at the bottom of this file.

## What was wrong

Opening a `.docx` does not go through `word/fromToJSON.js`. x2t reads `w:jc` and writes one byte into the binary that sdkjs loads in `Serialize2.js`.

In core `v9.4.0.129`, paragraph writing (`OOXML/Binary/Document/BinWriter/BinaryWriterD.cpp`) sent `jcBoth`, `jcDistribute`, and `jcThaiDistribute` as byte `3` (`align_Justify`). The three alignments were identical before sdkjs ever saw them. `fromToJSON.js` is a separate JSON import path and had the same collapse (`thaiDist` / `dist` became justify).

Saving goes the other way. sdkjs writes `pPr.Jc` as a byte. Unpatched x2t only maps bytes 0–3. Any other byte leaves `m_oJc` at its default, and the file is saved as `left`. Shipping the sdkjs change without rebuilding x2t turns `thaiDistribute` into left alignment on the next save.

Numbering `lvlJc` already wrote the raw `EJc` value and read it back with `SetValueFromByte`, so numbering was left alone. Table and row alignment only have left, center, and right.

## Binary contract

Bytes are the sdkjs constants in `common/commonDefines.js`. Byte `5` stays `align_CenterContinuous`.

| `w:jc` | sdkjs | binary byte | x2t `SetValueFromByte` | saved `w:jc` |
| --- | --- | --- | --- | --- |
| `both` | `align_Justify` (3) | 3 | 0 (`jcBoth`) | `both` |
| `distribute` | `align_Distributed` (4) | 4 | 2 (`jcDistribute`) | `distribute` |
| `thaiDistribute` | `align_ThaiDistributed` (6) | 6 | 9 (`jcThaiDistribute`) | `thaiDistribute` |

`right` / `left` / `center` stay bytes 0 / 1 / 2 and still save as `right` / `left` / `center` (`jcEnd` / `jcStart` / `jcCenter`).

Core change (also `core-patches/thai-distribute.patch`, applied to tag `v9.4.0.129`):

- `OOXML/Binary/Document/BinWriter/BinReaderWriterDefines.h` — `align_Distributed = 4`, `align_ThaiDistributed = 6`
- `OOXML/Binary/Document/BinWriter/BinaryWriterD.cpp` — paragraph `jc` writer
- `OOXML/Binary/Document/BinReader/BinaryReaderD.cpp` — paragraph `jc` reader

## What sdkjs does

- `align_ThaiDistributed = 6` is read and written on the paragraph `pPr.Jc` path only.
- JSON (`fromToJSON.js`), the API (`SetJc` / `GetJc`: `distribute`, `thaiDistribute`), HTML copy/paste (`text-justify: distribute` and `text-justify: inter-cluster`), and the macro recorder preserve the three values. DrawingML `dist` and `thaiDist` map to the same constants.
- Layout for `thaiDistribute` divides the extra width by `(clusters - 1) + counted spaces` and adds that amount to every cluster terminator and every counted space. A cluster is a base plus following Thai marks (U+0E31, U+0E34–U+0E3A, U+0E47–U+0E4E). The extra sits on the cluster terminator, including the last cluster, so the ink ends on the margin and no gap is inserted before a mark. The last line and a soft line break are not stretched. `distribute` uses the same gaps and stretches every line, including the last.
- `both` still puts extra into spaces when the line has spaces. Thai marks are no longer counted as separate letters, so a justified line with no spaces does not open a gap before a mark.
- Line breaking walks the paragraph, caches the result on the paragraph text, and sets a break after Thai words. `Intl.Segmenter("th", { granularity: "word" })` is used when a probe of `ดูการบ้าน` yields a first word `ดู`. Otherwise the bundled ICU dictionary (`word/Editor/Thai/thai-dict.js`, Unicode License) does longest-match segmentation. Unknown Thai breaks one grapheme cluster at a time. Latin text is not given new break points. A word-final tone mark can end the word; a mark is never a break by itself.

`word/Editor/Thai/thai-dict.js` and `thai-break.js` are concatenated into the word, cell, slide, and visio bundles, ahead of `Run.js`.

## Limitations compared with Word

- Each cluster gap and each counted space receives the same extra. Word can weight a space more than the gap between two clusters. Phrase spacing becomes even instead of a hole at each space; it can still differ from Word by that weight.
- The last line of a `thaiDistribute` paragraph is not stretched, matching Word's last-line behavior for this alignment. `distribute` does stretch the last line.
- `Intl.Segmenter` and the ICU list can disagree on rare words. The dictionary is maximal matching, not Word's exact breaker.
- A cluster that starts in one run and ends in the next is kept together for spacing. Segmentation sees the whole paragraph, including text inside hyperlinks.

## Build sdkjs (tag v9.4.0.129)

From `cursor/thai-distribute-9.4.0-703e`:

```bash
python3 build/build.py --product word
python3 build/build.py --product cell
python3 build/build.py --product slide
python3 build/build.py --product visio
```

Outputs:

```text
deploy/sdkjs/word/sdk-all.js
deploy/sdkjs/word/sdk-all-min.js
deploy/sdkjs/cell/sdk-all.js
deploy/sdkjs/cell/sdk-all-min.js
deploy/sdkjs/slide/sdk-all.js
deploy/sdkjs/slide/sdk-all-min.js
deploy/sdkjs/visio/sdk-all.js
deploy/sdkjs/visio/sdk-all-min.js
```

`sdk-all-min.js` holds `align_ThaiDistributed` and the macro recorder. `sdk-all.js` holds layout, segmentation, and the docx binary reader. Replace both for every editor you serve. Paragraph layout is shared, so cell, slide, and visio change even when the document is a Word file opened from the word editor.

## Build and install x2t

Replacing the JavaScript bundles does not preserve `thaiDistribute` on open, and saving through the stock `x2t` writes `left`.

`core-patches/thai-distribute.patch` is that commit against tag `v9.4.0.129`. Pushing it to [n00n0i/core](https://github.com/n00n0i/core) returned 403 (`Permission to n00n0i/core.git denied to cursor[bot]`). The fork is visible, and this credential can push [n00n0i/sdkjs](https://github.com/n00n0i/sdkjs), but it cannot write the core fork. Apply the patch from a checkout that can push:

```bash
git clone --branch v9.4.0.129 --depth 1 https://github.com/ONLYOFFICE/core.git
cd core
git checkout -b cursor/thai-distribute-x2t-9.4.0-703e
git apply /path/to/sdkjs/core-patches/thai-distribute.patch
git commit -am "Keep distribute and thaiDistribute distinct in x2t paragraph jc."
git push git@github.com:n00n0i/core.git v9.4.0.129:refs/heads/release/v9.4.0
git push -u git@github.com:n00n0i/core.git cursor/thai-distribute-x2t-9.4.0-703e
```

Open the pull request into `release/v9.4.0` on the fork. The fork's `master` is a later merge, so a pull request into `master` would mix this fix with the rest of the 9.4 history. The tag is not on the fork until that `release/v9.4.0` push.

x2t is the native converter built from `X2tConverter/build/Qt/X2tSLN.pro`. ONLYOFFICE builds it with [build_tools](https://github.com/ONLYOFFICE/build_tools) (`scripts/build_sln.py`) against a full core tree and its third-party libraries. Use the build_tools revision that produced Document Server 9.4.0-129, check out `cursor/thai-distribute-x2t-9.4.0-703e` (or apply the patch onto a core tree already at `v9.4.0.129`), and build the `x2t` target.

On the server the binary is:

```text
/var/www/onlyoffice/documentserver/server/FileConverter/bin/x2t
```

Back up that file, then replace it with the binary you built. This patch does not change library ABIs. If `ldd` on the new binary asks for sonames that are not already in that `bin` directory, copy those libraries beside it as well. Restart the converter (`supervisorctl restart all`, or `systemctl restart ds-converter ds-docservice`).

## Install sdkjs on Document Server 9.4.0-129

Confirm the package before copying files:

```bash
dpkg -l onlyoffice-documentserver | awk '/^ii/{print $3}'
# 9.4.0-129
```

Copy the eight bundles over:

```text
/var/www/onlyoffice/documentserver/sdkjs/word/sdk-all.js
/var/www/onlyoffice/documentserver/sdkjs/word/sdk-all-min.js
/var/www/onlyoffice/documentserver/sdkjs/cell/sdk-all.js
/var/www/onlyoffice/documentserver/sdkjs/cell/sdk-all-min.js
/var/www/onlyoffice/documentserver/sdkjs/slide/sdk-all.js
/var/www/onlyoffice/documentserver/sdkjs/slide/sdk-all-min.js
/var/www/onlyoffice/documentserver/sdkjs/visio/sdk-all.js
/var/www/onlyoffice/documentserver/sdkjs/visio/sdk-all-min.js
```

Then flush the server cache and reload the browser without a cached script:

```bash
documentserver-flush-cache.sh
supervisorctl restart all
```

Docker Hub image tag for this server line is `onlyoffice/documentserver:9.4.0` (the image tag does not include the package revision `-129`). Confirm the container with `dpkg -l onlyoffice-documentserver` before using it as the base.

```dockerfile
FROM onlyoffice/documentserver:9.4.0
COPY x2t /var/www/onlyoffice/documentserver/server/FileConverter/bin/x2t
COPY sdkjs/word/sdk-all.js /var/www/onlyoffice/documentserver/sdkjs/word/sdk-all.js
COPY sdkjs/word/sdk-all-min.js /var/www/onlyoffice/documentserver/sdkjs/word/sdk-all-min.js
COPY sdkjs/cell/sdk-all.js /var/www/onlyoffice/documentserver/sdkjs/cell/sdk-all.js
COPY sdkjs/cell/sdk-all-min.js /var/www/onlyoffice/documentserver/sdkjs/cell/sdk-all-min.js
COPY sdkjs/slide/sdk-all.js /var/www/onlyoffice/documentserver/sdkjs/slide/sdk-all.js
COPY sdkjs/slide/sdk-all-min.js /var/www/onlyoffice/documentserver/sdkjs/slide/sdk-all-min.js
COPY sdkjs/visio/sdk-all.js /var/www/onlyoffice/documentserver/sdkjs/visio/sdk-all.js
COPY sdkjs/visio/sdk-all-min.js /var/www/onlyoffice/documentserver/sdkjs/visio/sdk-all-min.js
RUN chmod 755 /var/www/onlyoffice/documentserver/server/FileConverter/bin/x2t
```

## Other release branches

`cursor/thai-distribute-9.4.0-703e` is already based on `v9.4.0.129`. Cherry-pick that commit onto another sdkjs release tag. Cherry-pick `cursor/thai-distribute-x2t-9.4.0-703e` onto the matching core tag, or `git apply core-patches/thai-distribute.patch` from a core checkout at `v9.4.0.129`. Rebuild both artifacts from those commits. Do not copy 9.4.0-129 bundles onto a different server version.

## Tests

`node tests/word/unit-tests/run-thai-distribute.js` and `tests/word/unit-tests/thai-distribute.html` check the three `w:jc` values (bytes 3, 4, and 6), cluster gaps, and that `ดูการบ้าน` breaks after `ดู`. On `cursor/thai-distribute-9.4.0-703e`, `tests/word/document-calculation/paragraph.html` also lays that string out in the editor and expects lines `ดู` and `การบ้าน`. `master` no longer has that harness, so the layout script is not wired there.

## Pull requests

- sdkjs: https://github.com/n00n0i/sdkjs/pull/1 (`cursor/thai-distribute-master-703e` into `master`). Deploy from `cursor/thai-distribute-9.4.0-703e`.
- core: not opened. The credential used here cannot push n00n0i/core, so there is no core pull request to link. Use `core-patches/thai-distribute.patch` and the commands above.
