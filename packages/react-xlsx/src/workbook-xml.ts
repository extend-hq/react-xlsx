type XmlTree = [string, Array<[string, string]>, Array<XmlTree | string>];
type XmlRecord = [string, XmlTree];

let decodeArchive: ((bytes: Uint8Array) => Record<string, Uint8Array>) | undefined;

export function configureWorkbookArchiveDecoder(decode: NonNullable<typeof decodeArchive>) { decodeArchive = decode; }
export function decodeWorkbookArchive(bytes: Uint8Array) {
  if (!decodeArchive) throw new Error("The workbook archive decoder is not initialized.");
  return decodeArchive(bytes);
}

let decodeXml: ((xml: string) => string) | undefined;
const trees = new Map<string, XmlTree>();
const owners = new Map<string, Set<symbol>>();
const ownedKeys = new Map<symbol, Set<string>>();
const XMLNS = "http://www.w3.org/2000/xmlns/";
const XML = "http://www.w3.org/XML/1998/namespace";

export function configureWorkbookXmlDecoder(decode: (xml: string) => string) {
  decodeXml = decode;
}

export function takeWorkbookXmlRecords(keep?: ReadonlySet<string>): XmlRecord[] {
  const records = Array.from(trees).filter(([xml]) => !keep || keep.has(xml));
  trees.clear();
  return records;
}

function retainTree(xml: string, tree: XmlTree, cacheOwners: ReadonlySet<symbol>) {
  trees.set(xml, tree);
  for (const owner of cacheOwners) {
    const keys = ownedKeys.get(owner);
    if (!keys) continue;
    keys.add(xml);
    let set = owners.get(xml);
    if (!set) owners.set(xml, set = new Set());
    set.add(owner);
  }
}

export function restoreWorkbookXmlRecords(records: XmlRecord[]): () => void {
  const owner = Symbol();
  ownedKeys.set(owner, new Set());
  for (const [xml, tree] of records) retainTree(xml, tree, new Set([owner]));
  return () => {
    for (const xml of ownedKeys.get(owner) ?? []) {
      const set = owners.get(xml);
      set?.delete(owner);
      if (!set?.size) { owners.delete(xml); trees.delete(xml); }
    }
    ownedKeys.delete(owner);
  };
}

class ParsedNode {
  readonly nodeType: number;
  private descendants?: { version: number; nodes: ParsedNode[] };
  parentNode: ParsedNode | null = null;
  ownerDocument: ParsedDocument;
  childNodes: ParsedNode[] = [];
  name: string;
  readonly localName: string;
  readonly prefix: string | null;
  data: string;
  attributeValues = new Map<string, string>();
  explicitNamespace: string | null | undefined;

  constructor(document: ParsedDocument, type: number, name = "", data = "") {
    this.ownerDocument = document;
    this.nodeType = type;
    this.name = name;
    const colon = name.indexOf(":");
    this.localName = colon < 0 ? name : name.slice(colon + 1);
    this.prefix = colon < 0 ? null : name.slice(0, colon);
    this.data = data;
  }

  get tagName() { return this.name; }
  get nodeName() { return this.name || (this.nodeType === 3 ? "#text" : "#document"); }
  get nodeValue() { return this.nodeType === 3 ? this.data : null; }
  get namespaceURI(): string | null { return this.explicitNamespace !== undefined ? this.explicitNamespace : this.lookupNamespaceURI(this.prefix); }
  get children() { return this.childNodes.filter(node => node.nodeType === 1); }
  get firstChild() { return this.childNodes[0] ?? null; }
  get firstElementChild() { return this.children[0] ?? null; }
  get parentElement(): ParsedNode | null { return this.parentNode?.nodeType === 1 ? this.parentNode : null; }
  get nextSibling(): ParsedNode | null { return this.parentNode?.childNodes[this.parentNode.childNodes.indexOf(this) + 1] ?? null; }
  get textContent(): string { return this.nodeType === 3 ? this.data : this.childNodes.map(node => node.textContent).join(""); }
  set textContent(text: string) {
    if (this.nodeType === 3) { this.data = text; return; }
    this.childNodes.forEach(node => { node.parentNode = null; });
    this.childNodes = [];
    this.ownerDocument.version++;
    if (text) this.appendChild(this.ownerDocument.createTextNode(text));
  }
  get attributes() {
    return Array.from(this.attributeValues, ([name, value]) => ({
      name, value, localName: name.split(":").at(-1),
      namespaceURI: name === "xmlns" || name.startsWith("xmlns:") ? XMLNS
        : name.includes(":") ? this.lookupNamespaceURI(name.split(":")[0]) : null
    }));
  }
  lookupNamespaceURI(prefix: string | null): string | null {
    if (prefix === "xml") return XML;
    if (prefix === "xmlns") return XMLNS;
    const key = prefix ? `xmlns:${prefix}` : "xmlns";
    return this.attributeValues.get(key) ?? this.parentNode?.lookupNamespaceURI(prefix) ?? null;
  }
  getAttribute(name: string) { return this.attributeValues.get(name) ?? null; }
  hasAttribute(name: string) { return this.attributeValues.has(name); }
  getAttributeNS(namespace: string | null, localName: string) {
    return this.attributes.find(attribute => attribute.namespaceURI === namespace && attribute.localName === localName)?.value ?? null;
  }
  setAttribute(name: string, value: string) { this.attributeValues.set(name, String(value)); }
  setAttributeNS(namespace: string | null, name: string, value: string) {
    this.setAttribute(name, value);
    const prefix = name.includes(":") ? name.split(":")[0] : null;
    if (namespace && namespace !== XMLNS && prefix && this.lookupNamespaceURI(prefix) !== namespace) {
      this.setAttribute(`xmlns:${prefix}`, namespace);
    }
  }
  removeAttribute(name: string) { this.attributeValues.delete(name); }
  appendChild<T extends ParsedNode>(node: T): T {
    node.remove(); node.parentNode = this; this.childNodes.push(node); this.ownerDocument.version++; return node;
  }
  insertBefore<T extends ParsedNode>(node: T, successor: ParsedNode | null): T {
    if (!successor) return this.appendChild(node);
    if (node === successor) return node;
    node.remove();
    const index = this.childNodes.indexOf(successor);
    if (index < 0) throw new Error("XML insertion target is not a child.");
    node.parentNode = this; this.childNodes.splice(index, 0, node); this.ownerDocument.version++; return node;
  }
  removeChild<T extends ParsedNode>(node: T): T {
    const index = this.childNodes.indexOf(node);
    if (index >= 0) { this.childNodes.splice(index, 1); node.parentNode = null; this.ownerDocument.version++; }
    return node;
  }
  remove() { this.parentNode?.removeChild(this); }
  cloneNode(deep = false): ParsedNode {
    const node = new ParsedNode(this.ownerDocument, this.nodeType, this.name, this.data);
    node.attributeValues = new Map(this.attributeValues);
    node.explicitNamespace = this.namespaceURI;
    for (const attribute of this.attributes) {
      if (attribute.namespaceURI && attribute.namespaceURI !== XMLNS && attribute.name.includes(":")) {
        node.attributeValues.set(`xmlns:${attribute.name.split(":")[0]}`, attribute.namespaceURI);
      }
    }
    if (deep) this.childNodes.forEach(child => node.appendChild(child.cloneNode(true)));
    return node;
  }
  getElementsByTagName(name: string): ParsedNode[] {
    if (this.descendants?.version === this.ownerDocument.version) return name === "*" ? this.descendants.nodes : this.descendants.nodes.filter(node => node.name === name);
    const nodes: ParsedNode[] = [];
    const visit = (node: ParsedNode) => {
      for (const child of node.childNodes) {
        if (child.nodeType !== 1) continue;
        nodes.push(child);
        visit(child);
      }
    };
    visit(this);
    this.descendants = { version: this.ownerDocument.version, nodes };
    return name === "*" ? nodes : nodes.filter(node => node.name === name);
  }
  getElementsByTagNameNS(namespace: string, localName: string) {
    return this.getElementsByTagName("*").filter(node => (namespace === "*" || node.namespaceURI === namespace) && (localName === "*" || node.localName === localName));
  }
  querySelector(selector: string) { return this.querySelectorAll(selector)[0] ?? null; }
  querySelectorAll(selector: string) {
    const names = selector.split(",").map(name => name.trim());
    if (names.some(name => !/^[\w:-]+$/.test(name))) throw new Error("Unsupported XML selector.");
    return this.getElementsByTagName("*").filter(node => names.includes(node.localName) || names.includes(node.name));
  }
}

class ParsedDocument extends ParsedNode {
  version = 0;
  cacheOwners = new Set<symbol>();
  constructor() { super(null as unknown as ParsedDocument, 9); this.ownerDocument = this; }
  get documentElement() { return this.firstElementChild; }
  implementation = {
    createDocument: (namespace: string | null, name: string) => {
      const document = new ParsedDocument();
      document.cacheOwners = new Set(this.cacheOwners);
      if (name) document.appendChild(document.createElementNS(namespace, name));
      return document;
    }
  };
  createElementNS(namespace: string | null, name: string) {
    const node = new ParsedNode(this, 1, name); node.explicitNamespace = namespace; return node;
  }
  createTextNode(text: string) { return new ParsedNode(this, 3, "", text); }
  importNode(node: ParsedNode, deep: boolean) {
    const copy = node.cloneNode(deep);
    const attach = (child: ParsedNode) => { child.ownerDocument = this; child.childNodes.forEach(attach); };
    attach(copy); return copy;
  }
}

function hydrate(tree: XmlTree, document: ParsedDocument, parent: ParsedNode) {
  for (const child of tree[2]) {
    if (typeof child === "string") { parent.appendChild(document.createTextNode(child)); continue; }
    const node = new ParsedNode(document, 1, child[0]);
    node.attributeValues = new Map(child[1]);
    parent.appendChild(node);
    hydrate(child, document, node);
  }
}

export function parseWorkbookXml(xml: string): Document {
  let tree = trees.get(xml);
  if (!tree) {
    if (!decodeXml) throw new Error("The workbook XML decoder is not initialized.");
    tree = JSON.parse(decodeXml(xml)) as XmlTree;
    trees.set(xml, tree);
  }
  const document = new ParsedDocument();
  document.cacheOwners = new Set(owners.get(xml));
  hydrate(tree, document, document);
  return document as unknown as Document;
}

export function serializeWorkbookXml(source: Node): string {
  const root = source as unknown as ParsedNode;
  const escape = (text: string) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const attribute = (text: string) => escape(text).replace(/"/g, "&quot;").replace(/\r/g, "&#13;").replace(/\n/g, "&#10;").replace(/\t/g, "&#9;");
  const serialize = (node: ParsedNode, inherited: Map<string, string>): string => {
    if (node.nodeType === 3) return escape(node.data);
    if (node.nodeType === 9) return node.childNodes.map(child => serialize(child, inherited)).join("");
    const attrs = new Map(node.attributeValues);
    const namespaces = new Map(inherited);
    for (const [key, value] of attrs) if (key === "xmlns" || key.startsWith("xmlns:")) namespaces.set(key, value);
    const namespace = node.namespaceURI;
    const declaration = node.prefix ? `xmlns:${node.prefix}` : "xmlns";
    if (namespace && namespaces.get(declaration) !== namespace) { attrs.set(declaration, namespace); namespaces.set(declaration, namespace); }
    if (!namespace && !node.prefix && namespaces.get("xmlns")) { attrs.set("xmlns", ""); namespaces.set("xmlns", ""); }
    for (const attr of node.attributes) {
      if (!attr.namespaceURI || attr.namespaceURI === XMLNS || attr.namespaceURI === XML || !attr.name.includes(":")) continue;
      const key = `xmlns:${attr.name.split(":")[0]}`;
      if (namespaces.get(key) !== attr.namespaceURI) { attrs.set(key, attr.namespaceURI); namespaces.set(key, attr.namespaceURI); }
    }
    const start = `<${node.name}${Array.from(attrs, ([key, value]) => ` ${key}="${attribute(value)}"`).join("")}`;
    node.attributeValues = attrs;
    return node.childNodes.length ? `${start}>${node.childNodes.map(child => serialize(child, namespaces)).join("")}</${node.name}>` : `${start}/>`;
  };
  const xml = serialize(root, new Map());
  const flatten = (node: ParsedNode): XmlTree | string => node.nodeType === 3 ? node.data : [node.name, Array.from(node.attributeValues), node.childNodes.map(flatten)];
  retainTree(xml, root.nodeType === 9 ? flatten(root) as XmlTree : ["", [], [flatten(root)]], root.ownerDocument.cacheOwners);
  return xml;
}

export function installWorkerXmlAdapters() {
  Object.assign(globalThis, {
    DOMParser: class { parseFromString(xml: string) { return parseWorkbookXml(xml); } },
    XMLSerializer: class { serializeToString(node: Node) { return serializeWorkbookXml(node); } },
    Node: { ELEMENT_NODE: 1, TEXT_NODE: 3, DOCUMENT_NODE: 9 }
  });
}
