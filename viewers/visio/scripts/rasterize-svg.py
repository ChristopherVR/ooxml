"""Secondary SVG test rendering with system librsvg/Cairo, not browser evidence.

Reads one bounded SVG on stdin and writes PNG on stdout. No network or document files.
Only used with the generated test scene in test-svg-rasterization.mjs.
"""
import ctypes as C
import ctypes.util
import math
import pathlib
import sys
import tempfile


def library(name):
    path = ctypes.util.find_library(name)
    if not path:
        raise RuntimeError(f"Secondary SVG test requires system {name}")
    return C.CDLL(path)


rsvg, cairo, gobject = library("rsvg-2"), library("cairo"), library("gobject-2.0")


def function(library, name, result, args):
    fn = getattr(library, name)
    fn.restype, fn.argtypes = result, args
    return fn


class Rectangle(C.Structure):
    _fields_ = [(name, C.c_double) for name in ["x", "y", "width", "height"]]


new_handle = function(rsvg, "rsvg_handle_new_from_data", C.c_void_p, [C.c_char_p, C.c_size_t, C.c_void_p])
set_dpi = function(rsvg, "rsvg_handle_set_dpi", None, [C.c_void_p, C.c_double])
size = function(rsvg, "rsvg_handle_get_intrinsic_size_in_pixels", C.c_int, [C.c_void_p, C.POINTER(C.c_double), C.POINTER(C.c_double)])
render = function(rsvg, "rsvg_handle_render_document", C.c_int, [C.c_void_p, C.c_void_p, C.POINTER(Rectangle), C.c_void_p])
surface_create = function(cairo, "cairo_image_surface_create", C.c_void_p, [C.c_int, C.c_int, C.c_int])
context_create = function(cairo, "cairo_create", C.c_void_p, [C.c_void_p])
write_png = function(cairo, "cairo_surface_write_to_png", C.c_int, [C.c_void_p, C.c_char_p])
context_destroy = function(cairo, "cairo_destroy", None, [C.c_void_p])
surface_destroy = function(cairo, "cairo_surface_destroy", None, [C.c_void_p])
unref = function(gobject, "g_object_unref", None, [C.c_void_p])

source = sys.stdin.buffer.read(16 * 1024 * 1024 + 1)
if len(source) > 16 * 1024 * 1024:
    raise RuntimeError("Secondary SVG test input exceeds 16 MiB")
handle = new_handle(source, len(source), None)
if not handle:
    raise RuntimeError("Secondary SVG renderer rejected XML")
surface = context = None
try:
    set_dpi(handle, 96)
    width, height = C.c_double(), C.c_double()
    if not size(handle, C.byref(width), C.byref(height)):
        raise RuntimeError("Secondary SVG renderer could not resolve dimensions")
    if not all(math.isfinite(n) and 0 < n <= 4096 for n in [width.value, height.value]):
        raise RuntimeError("Secondary SVG test dimensions exceed limit")
    surface = surface_create(0, math.ceil(width.value), math.ceil(height.value))
    context = context_create(surface)
    viewport = Rectangle(0, 0, width.value, height.value)
    if not render(handle, context, C.byref(viewport), None):
        raise RuntimeError("Secondary SVG rendering failed")
    with tempfile.TemporaryDirectory(prefix="visio-svg-pixels-") as temp:
        output = pathlib.Path(temp) / "pixels.png"
        if write_png(surface, str(output).encode()) != 0:
            raise RuntimeError("Secondary SVG test could not encode PNG")
        sys.stdout.buffer.write(output.read_bytes())
finally:
    if context:
        context_destroy(context)
    if surface:
        surface_destroy(surface)
    unref(handle)
