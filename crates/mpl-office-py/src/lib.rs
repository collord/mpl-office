//! PyO3 bindings for mpl-office-core.
//!
//! Exposes a low-level `convert_svg_to_drawingml` function and a
//! `ConvertOptions` class. The Python layer on top adds `python-pptx` /
//! `python-docx` injection helpers and a matplotlib backend.

use mpl_office_core::{
    convert_svg_to_drawingml as rust_convert,
    convert_svg_to_drawingml_with_images as rust_convert_with_images,
    ConvertOptions as RustOptions,
};
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyBytes;

#[pyclass(name = "ConvertOptions")]
#[derive(Clone)]
struct PyConvertOptions {
    #[pyo3(get, set)]
    source_dpi: f64,
    #[pyo3(get, set)]
    target_width_emu: Option<i64>,
    #[pyo3(get, set)]
    target_height_emu: Option<i64>,
    #[pyo3(get, set)]
    offset_x_emu: i64,
    #[pyo3(get, set)]
    offset_y_emu: i64,
}

#[pymethods]
impl PyConvertOptions {
    #[new]
    #[pyo3(signature = (source_dpi=96.0, target_width_emu=None, target_height_emu=None, offset_x_emu=0, offset_y_emu=0))]
    fn new(
        source_dpi: f64,
        target_width_emu: Option<i64>,
        target_height_emu: Option<i64>,
        offset_x_emu: i64,
        offset_y_emu: i64,
    ) -> Self {
        Self {
            source_dpi,
            target_width_emu,
            target_height_emu,
            offset_x_emu,
            offset_y_emu,
        }
    }

    fn __repr__(&self) -> String {
        format!(
            "ConvertOptions(source_dpi={}, target_width_emu={:?}, target_height_emu={:?}, offset_x_emu={}, offset_y_emu={})",
            self.source_dpi, self.target_width_emu, self.target_height_emu, self.offset_x_emu, self.offset_y_emu
        )
    }
}

impl From<&PyConvertOptions> for RustOptions {
    fn from(o: &PyConvertOptions) -> Self {
        RustOptions {
            source_dpi: o.source_dpi,
            target_width_emu: o.target_width_emu,
            target_height_emu: o.target_height_emu,
            offset_x_emu: o.offset_x_emu,
            offset_y_emu: o.offset_y_emu,
        }
    }
}

#[pyfunction]
#[pyo3(signature = (svg, options=None))]
fn convert_svg_to_drawingml(svg: &str, options: Option<&PyConvertOptions>) -> PyResult<String> {
    let default_opts = PyConvertOptions::new(96.0, None, None, 0, 0);
    let opts = options.unwrap_or(&default_opts);
    let rust_opts: RustOptions = opts.into();
    rust_convert(svg, &rust_opts).map_err(|e| PyValueError::new_err(format!("{}", e)))
}

/// Like [`convert_svg_to_drawingml`] but also returns the list of raster
/// images extracted from the SVG. Each entry is a
/// ``(sentinel, bytes, format)`` tuple — the caller is expected to
/// register ``bytes`` as an image part in the destination OOXML file,
/// obtain a real relationship id, and replace every occurrence of
/// ``sentinel`` in the XML (appearing as ``r:embed="{sentinel}"``) with
/// that id.
#[pyfunction]
#[pyo3(signature = (svg, options=None))]
fn convert_svg_to_drawingml_with_images<'py>(
    py: Python<'py>,
    svg: &str,
    options: Option<&PyConvertOptions>,
) -> PyResult<(String, Vec<(String, Bound<'py, PyBytes>, String)>)> {
    let default_opts = PyConvertOptions::new(96.0, None, None, 0, 0);
    let opts = options.unwrap_or(&default_opts);
    let rust_opts: RustOptions = opts.into();
    let (xml, images) = rust_convert_with_images(svg, &rust_opts)
        .map_err(|e| PyValueError::new_err(format!("{}", e)))?;
    let py_images = images
        .into_iter()
        .map(|img| {
            let data = PyBytes::new_bound(py, &img.bytes);
            (img.sentinel, data, img.format)
        })
        .collect();
    Ok((xml, py_images))
}

/// The native extension module. Exposed as `mpl_office._native` from Python.
#[pymodule]
fn _native(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(convert_svg_to_drawingml, m)?)?;
    m.add_function(wrap_pyfunction!(convert_svg_to_drawingml_with_images, m)?)?;
    m.add_class::<PyConvertOptions>()?;
    Ok(())
}
