import * as React from 'react';
import {createRoot} from 'react-dom/client';
import {setWasmSource, XlsxViewer, XlsxViewerProvider, useXlsxViewerController} from '@extend-ai/react-xlsx';
import {getSheetsWasmModule} from '@extend-ai/react-xlsx/internal';
setWasmSource("/duke_sheets_wasm_bg.wasm");
const wasm = await getSheetsWasmModule();
const params = new URLSearchParams(location.search);
const bytes = new Uint8Array(await (await fetch('/@fs/Users/andrewluo/react-xlsx/output/formula-parity/' + (params.get('file') || 'seed.xlsx'))).arrayBuffer());
window.parityCases = await (await fetch('/@fs/Users/andrewluo/react-xlsx/output/formula-parity/cases.json')).json();
function App() {
  const [canvas,setCanvas] = React.useState(new URLSearchParams(location.search).get('canvas') !== 'false');
  const controller = useXlsxViewerController({file:bytes.buffer,readOnly:params.has('readonly'),allowResizeInReadOnly:true,useWorker:params.has('worker')});
  window.checkController=controller;
  window.checkWasm=wasm;
  window.paritySetCanvas=setCanvas;
  return <><button onClick={()=>setCanvas(!canvas)}>Switch renderer</button><span>{canvas?'canvas':'DOM'}</span><div style={{height:650,width:'min(1050px, 100vw)'}}><XlsxViewerProvider controller={controller}><XlsxViewer experimentalCanvas={canvas} allowResizeInReadOnly /></XlsxViewerProvider></div></>;
}
createRoot(document.getElementById('root')!).render(<App/>);
