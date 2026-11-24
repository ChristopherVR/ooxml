# JBIG2 fixtures

`jbig2.json` holds JBIG2 streams and the pages the jbig2dec reference decoder (0.20, used only
to compute the expected pixels, never shipped or linked) makes of them: packed rows with 1 for
black, or their SHA-256 for large pages. `generate.py` writes it; run it in Debian trixie, which
packages both tools:

```sh
docker run --rm -v "$PWD:/work" -w /work debian:trixie bash -c \
  "apt-get update -qq && apt-get install -y -qq jbig2 jbig2dec python3-pil && python3 generate.py"
```

| Case              | Source                                                                                         |
| ----------------- | ---------------------------------------------------------------------------------------------- |
| `generic`         | jbig2enc 0.30 (Apache-2.0), `jbig2 -p`: generic region, template 0, PDF embedded stream       |
| `generic tpgdon`  | jbig2enc 0.30, `jbig2 -p -d`: the same with typical prediction (TPGDON)                        |
| `symbol`          | jbig2enc 0.30, `jbig2 -s -p`: symbol dictionary in the globals stream, text region in the page |
| `file`            | jbig2enc 0.30, `jbig2`: a whole file in the sequential organisation                            |
| `annex h`         | `annex-h.jbig2`, see below                                                                     |
| `pdfbox issue 21` | `21.glob` and `21.jb2`, see below                                                              |

The images jbig2enc encodes are drawn by `generate.py` with Pillow.

- `annex-h.jbig2` is the three-page test bitstream of ITU-T T.88 Annex H.2 (Huffman and
  arithmetic symbol dictionaries, refinement and aggregate coding, text regions with
  refinement, MMR and arithmetic generic regions with TPGDON, pattern dictionaries and halftone
  regions). The bytes are those published in the Recommendation; this copy was taken from
  `annex-h.jbig2` in https://github.com/ArtifexSoftware/jbig2dec (commit
  `5df2b837e72d1d781fcc29c0d21f82222c2e7af4`, "Add the three-page test file from appendix h of
  the jbig2 spec"), unchanged.
- `21.glob` and `21.jb2` are the globals and page streams of a PDF image (Huffman coded symbol
  dictionaries and text regions, 2560 x 3296 pixels) from
  https://github.com/apache/pdfbox-jbig2 (Apache License 2.0),
  `src/test/resources/org/apache/pdfbox/jbig2/github/21.glob` and `21.jb2` (repository commit
  `2d62947a279423dfdadcd4ca2d894c55b4e665d9`), unchanged.

The arithmetic coder test sequence of T.88 Annex H.2 is spelled out in `../../jbig2-arith.test.ts`.
