use rust_office_art::{PictureContext, ResolvedColor};
use wasm_bindgen::prelude::*;

struct Context(js_sys::Function);

#[derive(serde::Deserialize)]
struct Color {
    color: String,
    opacity: Option<f64>,
}

impl PictureContext for Context {
    fn resolve_color(&self, xml: &str) -> Option<ResolvedColor> {
        let value = self
            .0
            .call1(&JsValue::NULL, &JsValue::from_str(xml))
            .ok()?
            .as_string()?;
        let color: Color = serde_json::from_str(&value).ok()?;
        Some(ResolvedColor {
            color: color.color,
            opacity: color.opacity,
        })
    }
}

#[wasm_bindgen(js_name = renderPicture)]
pub fn render_picture(
    xml: &str,
    source: &str,
    width: f64,
    height: f64,
    colors: js_sys::Function,
) -> Result<Option<String>, JsError> {
    rust_office_art::render_picture(xml, source, width, height, &Context(colors))
        .map(|picture| serde_json::to_string(&picture).map_err(JsError::from))
        .transpose()
}

#[wasm_bindgen(js_name = normalizeRaster)]
pub fn normalize_raster(source: &str) -> Option<String> {
    rust_office_art::normalize_raster(source)
}
