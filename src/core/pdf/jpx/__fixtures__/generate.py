"""
Generates jpx.json: small JPEG 2000 files made by public encoders, with their source pixels and,
for the lossy ones, the reference decode of OpenJPEG's opj_decompress.

Needs Python 3 with Pillow (built with OpenJPEG) and the OpenJPEG command line tools
(https://github.com/uclouvain/openjpeg/releases, v2.5.4 was used). Run from this folder:

    OPJ_BIN=path/to/openjpeg/bin python generate.py
"""

import base64
import json
import os
import random
import subprocess
import tempfile

from PIL import Image

BIN = os.environ.get("OPJ_BIN", "")
COMPRESS = os.path.join(BIN, "opj_compress")
DECOMPRESS = os.path.join(BIN, "opj_decompress")
W, H = 33, 17
rng = random.Random(319)


def smooth(depth, channel, signed=False):
    """A gradient with some texture, so that lossy coding keeps it recognisable."""
    top = (1 << depth) - 1
    out = []
    for y in range(H):
        for x in range(W):
            ramp = (x * 6 + y * 9 + channel * 50) % 360
            base = 30 + (ramp if ramp < 180 else 360 - ramp) + rng.randrange(0, 8)
            value = base * top // 255
            out.append(value - (1 << (depth - 1)) if signed else value)
    return out


SOURCES = {
    "gray8": {"components": [smooth(8, 0)], "depth": 8},
    "rgb8": {"components": [smooth(8, c) for c in range(3)], "depth": 8},
    "rgba8": {"components": [smooth(8, c) for c in range(4)], "depth": 8},
    "gray16": {"components": [smooth(16, 0)], "depth": 16},
    "gray12": {"components": [smooth(12, 0)], "depth": 12},
    "signed8": {"components": [smooth(8, 0, True)], "depth": 8, "signed": True},
    "yuv420": {"components": [smooth(8, c) for c in range(3)], "depth": 8, "sub": [1, 2, 2]},
}


def subsample(values, factor):
    w, h = -(-W // factor), -(-H // factor)
    return [values[min(H - 1, y * factor) * W + min(W - 1, x * factor)] for y in range(h) for x in range(w)]


for source in SOURCES.values():
    if "sub" in source:
        source["components"] = [subsample(c, f) for c, f in zip(source["components"], source["sub"])]
        # opj_compress reads floor(W * H / (dx * dy)) samples of a sub-sampled raw plane and
        # leaves the rest 0, so the source says the same.
        for c, f in zip(source["components"], source["sub"]):
            for i in range(W * H // (f * f), len(c)):
                c[i] = 0


def write_input(name, path):
    source = SOURCES[name]
    comps, depth = source["components"], source["depth"]
    if name in ("yuv420", "signed8"):
        sub = source.get("sub", [1] * len(comps))
        data = b"".join(bytes(v & 0xFF for v in c[: W * H // (f * f)]) for c, f in zip(comps, sub))
        with open(path + ".raw", "wb") as f:
            f.write(data)
        layout = "u@1x1:2x2:2x2" if name == "yuv420" else "s"
        return path + ".raw", ["-F", f"{W},{H},{len(comps)},8,{layout}"]
    if name == "rgba8":
        image = Image.new("RGBA", (W, H))
        image.putdata(list(zip(*comps)))
        image.save(path + ".png")
        return path + ".png", []
    if depth > 8:
        header = f"P5\n{W} {H}\n{(1 << depth) - 1}\n".encode()
        data = b"".join(v.to_bytes(2, "big") for v in comps[0])
        with open(path + ".pgm", "wb") as f:
            f.write(header + data)
        return path + ".pgm", []
    magic, ext = ("P5", ".pgm") if len(comps) == 1 else ("P6", ".ppm")
    data = bytes(v for pixel in zip(*comps) for v in pixel)
    with open(path + ext, "wb") as f:
        f.write(f"{magic}\n{W} {H}\n255\n".encode() + data)
    return path + ext, []


# name, source, opj_compress arguments, extension, lossy
CASES = [
    ("gray-53", "gray8", ["-n", "3"], "j2k", False),
    ("rgb-53-tiles", "rgb8", ["-t", "16,8", "-n", "3", "-b", "8,8"], "jp2", False),
    ("rgb-97", "rgb8", ["-I", "-r", "3", "-n", "3"], "jp2", True),
    ("gray-97-layers-rpcl", "gray8", ["-I", "-r", "12,5,1.5", "-n", "4", "-p", "RPCL", "-c", "[16,16],[8,8]"], "j2k", True),
    ("gray-all-modes", "gray8", ["-M", "63", "-n", "3", "-b", "8,8", "-r", "10,4,1"], "j2k", False),
    ("gray-bypass-layers", "gray8", ["-M", "1", "-r", "20,5,1", "-b", "16,16", "-n", "3"], "j2k", False),
    ("gray-termall-reset", "gray8", ["-M", "6", "-r", "8,1", "-n", "3"], "j2k", False),
    ("rgb-vcausal-segsym-97", "rgb8", ["-M", "40", "-I", "-r", "4", "-n", "3"], "j2k", True),
    ("rgb-rlcp-layers", "rgb8", ["-p", "RLCP", "-r", "10,1", "-c", "[16,16]", "-n", "3"], "j2k", False),
    ("rgb-pcrl-tiles", "rgb8", ["-p", "PCRL", "-c", "[8,8]", "-t", "20,10", "-n", "3"], "j2k", False),
    ("yuv420-cprl", "yuv420", ["-mct", "0", "-p", "CPRL", "-c", "[8,8],[4,4]", "-n", "3"], "j2k", False),
    ("yuv420-rpcl-tiles", "yuv420", ["-mct", "0", "-p", "RPCL", "-c", "[8,8]", "-t", "16,16", "-n", "3"], "j2k", False),
    ("rgb-markers", "rgb8", ["-SOP", "-EPH", "-PLT", "-TLM", "-TP", "R", "-p", "RPCL", "-r", "20,1", "-n", "3"], "j2k", False),
    ("rgb-poc", "rgb8", ["-r", "10,1", "-n", "3", "-POC", "T0=0,0,1,2,3,CPRL/T0=0,0,2,3,3,LRCP"], "j2k", False),
    ("gray-roi", "gray8", ["-ROI", "c=0,U=4", "-n", "3"], "j2k", False),
    ("rgb-offsets", "rgb8", ["-d", "5,3", "-T", "2,1", "-t", "16,8", "-n", "3"], "j2k", False),
    ("gray16", "gray16", ["-n", "3"], "jp2", False),
    ("gray12", "gray12", ["-n", "2", "-b", "4,4"], "j2k", False),
    ("signed8", "signed8", ["-n", "3"], "j2k", False),
    ("rgba", "rgba8", ["-n", "3"], "jp2", False),
]


def read_pnm(path, components):
    """The samples of an 8-bit PGM or PPM (which may hold comment lines) written by opj_decompress."""
    with open(path, "rb") as f:
        data = f.read()
    return data[len(data) - W * H * components :]


def main():
    out = {"width": W, "height": H, "sources": {}, "cases": []}
    for name, source in SOURCES.items():
        depth = source["depth"]
        packed = b"".join(
            (v & ((1 << 16) - 1)).to_bytes(2, "big") if depth > 8 else bytes([v & 0xFF])
            for c in source["components"]
            for v in c
        )
        out["sources"][name] = {
            "depth": depth,
            "signed": source.get("signed", False),
            "sub": source.get("sub", [1] * len(source["components"])),
            "data": base64.b64encode(packed).decode(),
        }
    with tempfile.TemporaryDirectory() as tmp:
        for name, source, args, ext, lossy in CASES:
            input_path, extra = write_input(source, os.path.join(tmp, name))
            output = os.path.join(tmp, f"{name}.{ext}")
            run = subprocess.run([COMPRESS, "-i", input_path, "-o", output, *extra, *args], capture_output=True, text=True)
            if run.returncode:
                raise SystemExit(f"{name}: {run.stdout}{run.stderr}")
            case = {"name": name, "source": source, "tool": "opj_compress " + " ".join(args), "lossy": lossy}
            with open(output, "rb") as f:
                case["data"] = base64.b64encode(f.read()).decode()
            if lossy:
                pnm = os.path.join(tmp, f"{name}-ref.{'pgm' if source == 'gray8' else 'ppm'}")
                subprocess.run([DECOMPRESS, "-i", output, "-o", pnm], check=True, capture_output=True)
                count = 1 if source == "gray8" else 3
                case["reference"] = base64.b64encode(read_pnm(pnm, count)).decode()
            out["cases"].append(case)
        # Pillow (OpenJPEG through its own encoder settings).
        rgb = Image.new("RGB", (W, H))
        rgb.putdata(list(zip(*SOURCES["rgb8"]["components"])))
        gray = Image.new("L", (W, H))
        gray.putdata(SOURCES["gray8"]["components"][0])
        pillow = [
            ("pillow-rgb-97", "rgb8", rgb, {"irreversible": True, "quality_mode": "rates", "quality_layers": [4], "tile_size": (16, 16), "num_resolutions": 3}, True),
            ("pillow-gray-53", "gray8", gray, {"progression": "PCRL", "codeblock_size": (4, 4), "num_resolutions": 3}, False),
        ]
        for name, source, image, options, lossy in pillow:
            output = os.path.join(tmp, f"{name}.jp2")
            image.save(output, "JPEG2000", **options)
            case = {"name": name, "source": source, "tool": "Pillow " + repr(options), "lossy": lossy}
            with open(output, "rb") as f:
                case["data"] = base64.b64encode(f.read()).decode()
            if lossy:
                case["reference"] = base64.b64encode(Image.open(output).convert(image.mode).tobytes()).decode()
            out["cases"].append(case)
    with open("jpx.json", "w") as f:
        json.dump(out, f, indent=1)
        f.write("\n")


main()
