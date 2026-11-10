import * as React from "react";
import { createRoot } from "react-dom/client";
import { XlsxViewer, useXlsxViewerController, setWasmSource } from "@extend-ai/react-xlsx";
import { createXlsxViewer } from "@extend-ai/react-xlsx/vanilla";
import { FormulaEditorDemo } from "./src/components/formula-editor-demo";
import "./src/index.css";
setWasmSource("/duke_sheets_wasm_bg.wasm");
const params = new URLSearchParams(location.search);
const api = window as any;
const options = {src:"/editor-check.xlsx",height:440,readOnly:false,useWorker:false,showDefaultToolbar:false,experimentalCanvas:params.get("canvas")!=="false", enableCanvasSelectionAnimation:false};
if (params.get("mode") === "vanilla") {
  const viewer = createXlsxViewer(document.getElementById("root")!, {
    ...options,
    cellEditorStyle: { boxShadow:"inset 0 0 0 2px rgb(220, 38, 38)" },
    renderCellEditorOverlay: params.get("enabled") === "false" ? undefined : session => {
      api.session = session;
      const root = document.createElement("div");
      root.className="check-actions";
      Object.assign(root.style,{position:"absolute",top:"calc(100% + 8px)",pointerEvents:"auto"});
      const actions = session.isPreview ? [["Keep", () => session.commit()],["Discard", () => session.discardPreview()]] : [
        ["Preview",() => session.preview("=SUM(10,20,30)")],
        ["Apply",() => session.commit("=SUM(10,20,30)")],
        ["Replace",() => session.setValue("replacement")],
        ["Cancel",() => session.cancel()]
      ];
      for (const [label,callback] of actions) {const button=document.createElement("button");button.textContent=label as string;button.onclick=callback as any;root.append(button);}
      const input=document.createElement("input");input.placeholder="Overlay input";root.append(input);
      return root;
    }
  });
  api.viewer=viewer;
  api.setOptions=(next:any)=>viewer.setOptions(next);
  viewer.subscribe(controller=>{api.controller=controller;});
} else {
  function Demo() {
    const controller=useXlsxViewerController(options);
    const [extra,setExtra]=React.useState({});
    api.controller=controller;
    api.setOptions=(next:any)=>setExtra(previous=>({...previous,...next}));
    return <XlsxViewer {...options} controller={controller} renderCellEditorOverlay={params.get("enabled")==="false"?undefined:session=>{api.session=session;return <FormulaEditorDemo session={session} isDark={params.get("dark")==="true"}/>;}} {...extra}/>;
  }
  const root=createRoot(document.getElementById("root")!);
  api.viewer={destroy:()=>root.unmount()};
  root.render(<React.StrictMode><Demo/></React.StrictMode>);
}
