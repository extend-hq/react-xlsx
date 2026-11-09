import * as React from "react";
import { englishMessages, type XlsxViewerMessages } from "./messages";

export const XlsxMessagesContext = React.createContext<XlsxViewerMessages>(englishMessages);

export function useXlsxMessages() {
  return React.useContext(XlsxMessagesContext);
}
