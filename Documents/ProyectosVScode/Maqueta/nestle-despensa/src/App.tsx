// ==== INICIO CÓDIGO APP NESTLÉ (TypeScript) ====
import React, {
  useMemo,
  useRef,
  useState,
  useDeferredValue,
  useEffect,
} from "react";

declare global {
  interface Window {
    webkitSpeechRecognition?: any;
    SpeechRecognition?: any;
  }
}

import * as XLSX from "xlsx";

// ===== Tipos =====
export type Product = {
  sku: string;
  name: string;
  brand: string;
  type: string;
  price: number;
  oldPrice?: number;
  tags?: string[];
  image?: string;
  badges?: string[];
  description?: string;
};

type CartItem = { sku: string; qty: number };

type NormalizedRow = Record<string, unknown>;

// ===== Helpers =====
function badgeClass(label: string): string {
  const map: Record<string, string> = {
    Nuevo: "bg-sky-50 border-sky-200 text-sky-700",
    "Sin lactosa": "bg-emerald-50 border-emerald-200 text-emerald-700",
    Promo: "bg-rose-50 border-rose-200 text-rose-700",
    Oferta: "bg-amber-50 border-amber-200 text-amber-700",
  };
  return map[label] || "bg-slate-50 border-slate-200 text-slate-700";
}

function countBy(arr: string[]): Record<string, number> {
  const m: Record<string, number> = {};
  for (const x of arr) m[x] = (m[x] || 0) + 1;
  return m;
}

function parseTags(v: unknown): string[] {
  if (!v) return [];
  if (Array.isArray(v)) return v.map(String);
  return String(v)
    .split(/[;,|]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

function normalizeKey(k: unknown): string {
  return String(k || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[ _]+/g, " ")
    .trim();
}

function cleanNumberLike(input: unknown): string {
  let s = String(input ?? "");
  let out = "";
  for (const ch of s) {
    if ("0123456789,.,-$ ".includes(ch)) out += ch;
  }
  out = out.split("$").join("").split(" ").join("");
  out = out.replace(/\.(?=.*\.)/g, ""); // deja solo el último punto
  out = out.replace(",", ".");
  return out;
}

function isHttpUrl(u?: string): boolean {
  if (!u) return false;
  try {
    const p = new URL(u);
    return p.protocol === "http:" || p.protocol === "https:";
  } catch {
    return false;
  }
}

const FALLBACK_IMAGE =
  "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0nNTEyJyBoZWlnaHQ9JzUxMicgdmlld0JveD0nMCAwIDUxMiA1MTInIHhtbG5zPSdodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2Zyc+PHJlY3Qgd2lkdGg9JzUxMicgaGVpZ2h0PSc1MTInIGZpbGw9J3doaXRlJy8+PGNpcmNsZSBjeD0nMjU2JyBjeT0nMjU2JyByPScxOTAnIGZpbGw9JyNmZWY4ZTgnIHN0cm9rZT0nI2U1ZTVlNScgc3Ryb2tlLXdpZHRoPScxNScvPjx0ZXh0IHg9JzUwJScgeT0nNTAuNSUnIGR5PScuMzVlbW0nIGZvbnQtc2l6ZT0nNDAnIHRleHQtYW5jaG9yPSdtaWRkbGUnIGZpbGw9JyNFNDFDMjMnPk5lc3Rsw6kgQkIyQzwvdGV4dD48L3N2Zz4=";



const DEFAULT_PRODUCTS: Product[] = [
  {
    sku: "1001",
    name: "Leche Condensada Nestlé 397g",
    brand: "Nestlé",
    type: "Cremas y dulces lácteos",
    price: 1590,
    oldPrice: 1990,
    tags: ["leche", "postres", "recetas", "Cremas y dulces lácteos"],
    badges: ["Oferta"],
    image: "",
    description: "Leche condensada para postres y recetas.",
  },
  {
    sku: "1002",
    name: "Cereal Chocapic 500g",
    brand: "Nestlé",
    type: "Cereales",
    price: 3490,
    tags: ["desayuno", "chocolate", "Cereales"],
    badges: ["Nuevo"],
    image: "",
    description: "Cereal sabor chocolate.",
  },
  {
    sku: "1003",
    name: "Café Dolca 170g",
    brand: "Nestlé",
    type: "Café",
    price: 3490,
    oldPrice: 3990,
    tags: ["café", "instantáneo", "Café"],
    badges: ["Oferta"],
    image: "",
    description: "Café instantáneo clásico.",
  },
];

// === Normalizador ===
function normalizeIncoming(
  raw: NormalizedRow,
  options: { ignoreShowFlag: boolean } = { ignoreShowFlag: false }
): Product | null {
  const nk = (s: unknown) => normalizeKey(s);
  const getExact = (keys: string[]): unknown => {
    for (const k of Object.keys(raw || {})) {
      const left = nk(k);
      for (const want of keys) {
        if (left === nk(want)) return raw[k];
      }
    }
    return undefined;
  };

  const skuRaw = getExact(["SKU", "sku", "codigo", "código", "id", "ean", "upc"]);
  const nameRaw = getExact(["Producto", "producto", "name", "nombre", "titulo", "título"]);
  const brandRaw = getExact(["Marca", "marca", "brand"]);
  const typeRaw = getExact(["Categoría", "categoria", "categoría", "type"]);
  const deptRaw = getExact(["Departamento", "departamento"]);
  const topCatRaw = getExact(["Top categoría", "top categoria", "top categoría"]);
  const showRaw = getExact([
    "Mostrar en Tienda",
    "mostrar en tienda",
    "visible",
    "publicado",
    "activo",
    "online",
  ]);
  const priceRaw = getExact([
    "Precio Base",
    "price",
    "precio",
    "precio base",
    "precio normal",
    "precio final",
  ]);
  const oldRaw = getExact([
    "Precio normal",
    "oldPrice",
    "precio referencia",
    "precio anterior",
    "price_old",
  ]);
  const descRaw = getExact([
    "Descripción enriquecida",
    "descripcion",
    "descripción",
    "detalle",
    "ficha",
    "description",
  ]);
  const tagsRaw = getExact(["Tags", "tags", "etiquetas", "keywords", "palabras clave"]);
  const imgRaw = getExact(["URL_Imagen", "image", "imagen", "img", "foto", "url imagen", "url"]);

  const eanRaw = getExact(["EAN", "ean"]);
  const sapRaw = getExact(["SAP", "sap"]);
  const imgQueryRaw = getExact(["Imagen_Query", "imagen_query", "img_query"]);

  // Visibilidad por "Mostrar en Tienda"
  if (!options.ignoreShowFlag) {
    if (showRaw !== undefined) {
      const s = String(showRaw ?? "").toLowerCase().trim();
      const ok = ["si", "sí", "yes", "true", "1", "x", "ok"].includes(s);
      if (!ok) return null;
    }
  }

  const name = String(nameRaw ?? "").trim();
  if (!name) return null;

  let sku = String(skuRaw ?? "").trim();
  if (!sku) sku = `tmp-${Math.random().toString(36).slice(2, 9)}`;

  const brand = String(brandRaw ?? "").trim();
  let type = String(typeRaw ?? "").trim();
  if (!type) type = String(topCatRaw ?? deptRaw ?? "").trim();

  const toNum = (v: unknown): number => {
    const n = Number(cleanNumberLike(v));
    return Number.isFinite(n) ? n : 0;
  };

  let price = toNum(priceRaw);
let oldPrice: number | undefined = toNum(oldRaw);
if (price <= 0 && oldPrice > 0) {
  price = oldPrice;
  oldPrice = 0;
}
if (oldPrice <= 0) oldPrice = undefined;


  const tags = parseTags(tagsRaw);
  const addTag = (t: unknown) => {
    const s = String(t || "").trim();
    if (s && !tags.includes(s)) tags.push(s);
  };
  if (type) addTag(type);
  if (brand) addTag(brand);
  if (deptRaw) addTag(String(deptRaw));
  if (topCatRaw) addTag(String(topCatRaw));
  if (eanRaw) addTag(`ean:${String(eanRaw).trim()}`);
  if (sapRaw) addTag(`sap:${String(sapRaw).trim()}`);
  if (imgQueryRaw) addTag(`image_query:${String(imgQueryRaw).trim()}`);

  const image = String(imgRaw ?? "").trim();
  const description = String(descRaw ?? "").trim();

  const badges: string[] = [];
  if (tags.some((t) => normalizeKey(t) === normalizeKey("Sin lactosa"))) {
    badges.push("Sin lactosa");
  }
  if (typeof oldPrice === "number" && typeof price === "number" && oldPrice > price && price > 0) {
    badges.push("Oferta");
  }

  return { sku, name, brand, type, price, oldPrice, tags, image, badges, description };
}

// === Detección de cabeceras con verificación de datos debajo ===
function detectTableWithData(ws: XLSX.WorkSheet): {
  headers: string[];
  rows: NormalizedRow[];
  reason: string | null;
} {
  const norm = (s: unknown) =>
    normalizeKey(String(s || "").replace(/\./g, " ").replace(/-/g, " "));
  const rowsRaw = (XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }) as unknown[][]) || [];
  const isEmptyRow = (r: unknown[]): boolean =>
    !Array.isArray(r) ||
    r.length === 0 ||
    r.every((c) => String(c ?? "").trim() === "");
  const rows = rowsRaw.filter((r) => !isEmptyRow(r));

  const EXPECTED = new Set<string>([
    "sku",
    "ean",
    "sap",
    "producto",
    "sku activo",
    "producto activo",
    "mostrar en tienda",
    "descripcion",
    "descripción",
    "departamento",
    "categoria",
    "categoría",
    "marca",
    "precio base",
    "packs",
    "sin lactosa",
    "yoghurt y postre",
    "mundo bebe",
    "cafe",
    "café",
    "cereales",
    "chocolates",
    "galletas",
    "helados",
    "cremas y dulces lacteos",
    "maggi",
    "alimentos para mascotas",
    "nestle profesional",
    "saborizante para la leche",
    "nutricion adulto",
    "limpieza y aseo",
    "pilas y baterias",
    "bebidas y jugos",
    "top categoria",
    "top categoría",
    "tags",
    "descripcion enriquecida",
    "descripción enriquecida",
    "imagen query",
    "url imagen",
    "url_imagen",
    "url imagen",
  ]);

  const LIMIT_HEADER_SCAN = Math.min(rows.length, 200);
  let best = { idx: -1, headers: [] as string[], dataCount: -1, reason: "" };

  for (let i = 0; i < LIMIT_HEADER_SCAN; i++) {
    const candidate = rows[i] || [];
    const normalized = candidate.map(norm);
    const hitCount = normalized.filter((n) => EXPECTED.has(n)).length;

    if (hitCount >= 3) {
      const headers = candidate.map(
        (h, idx) => String(h || "").trim() || `col_${idx + 1}`
      );

      const lookaheadEnd = Math.min(rows.length, i + 1 + 100);
      let dataCount = 0;
      for (let r = i + 1; r < lookaheadEnd; r++) {
        const row = rows[r] || [];
        const someData = headers.some(
          (_, colIndex) => String((row as unknown[])[colIndex] ?? "").trim() !== ""
        );
        if (someData) dataCount++;
      }

      if (dataCount > best.dataCount) {
        best = { idx: i, headers, dataCount, reason: "" };
      }
    }
  }

  if (best.idx < 0) {
    return {
      headers: [],
      rows: [],
      reason: "No se encontró una fila de cabeceras con ≥3 campos esperados.",
    };
  }
  if (best.dataCount <= 0) {
    return {
      headers: best.headers,
      rows: [],
      reason: "Se detectó cabecera, pero no se observan filas con datos debajo.",
    };
  }

  const dataRows = rows.slice(best.idx + 1);
  const objects: NormalizedRow[] = dataRows
    .map((r) => {
      const obj: NormalizedRow = {};
      best.headers.forEach((h, i) => {
        (obj as any)[h] = (r as unknown[])[i];
      });
      return obj;
    })
    .filter((o) =>
      Object.values(o).some((v) => String(v ?? "").trim() !== "")
    );

  return {
    headers: best.headers,
    rows: objects,
    reason: objects.length ? null : "Sin filas luego de la cabecera.",
  };
}

// ===== Subcomponentes =====
interface CatalogViewProps {
  facetBrands: Record<string, number>;
  facetTypes: Record<string, number>;
  sorted: Product[];
  queryInput: string;
  setQueryInput: (v: string) => void;
  brandFilter: string | null;
  setBrandFilter: (v: string | null) => void;
  typeFilter: string | null;
  setTypeFilter: (v: string | null) => void;
  sortBy: string;
  setSortBy: (v: string) => void;
  favorites: Set<string>;
  onToggleFavorite: (sku: string) => void;
  onOpenDetail: (p: Product) => void;
  onAddToCart: (p: Product) => void;
}

function CatalogView({
  facetBrands,
  facetTypes,
  sorted,
  queryInput,
  setQueryInput,
  brandFilter,
  setBrandFilter,
  typeFilter,
  setTypeFilter,
  sortBy,
  setSortBy,
  favorites,
  onToggleFavorite,
  onOpenDetail,
  onAddToCart,
}: CatalogViewProps) {
  const searchRef = useRef<HTMLInputElement | null>(null);

  const brandOptions = useMemo(() => {
    const keys = Object.keys(facetBrands).filter(
      (k) => facetBrands[k] > 0 || k === (brandFilter || "")
    );
    return keys.sort((a, b) => a.localeCompare(b, "es"));
  }, [facetBrands, brandFilter]);

  const typeOptions = useMemo(() => {
    const keys = Object.keys(facetTypes).filter(
      (k) => facetTypes[k] > 0 || k === (typeFilter || "")
    );
    return keys.sort((a, b) => a.localeCompare(b, "es"));
  }, [facetTypes, typeFilter]);

  return (
    <>
      <div className="px-4 pb-2 mt-1">
        <div className="relative">
          <input
            ref={searchRef}
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 pr-9 shadow-sm focus:outline-none focus:ring-2 focus:ring-rose-200"
            placeholder="Buscar productos Nestlé…"
            value={queryInput}
            onChange={(e) => setQueryInput(e.target.value)}
            autoComplete="off"
            spellCheck={false}
          />
          {queryInput && (
            <button
              className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
              onClick={() => {
                setQueryInput("");
                searchRef.current?.focus();
              }}
              aria-label="Limpiar"
            >
              ✕
            </button>
          )}
        </div>
      </div>

      <div className="px-4 pb-2 grid grid-cols-2 gap-3">
        <div>
          <label className="block text-xs text-slate-600 mb-1">Marca</label>
          <select
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 shadow-sm"
            value={brandFilter || "__all__"}
            onChange={(e) =>
              setBrandFilter(e.target.value === "__all__" ? null : e.target.value)
            }
          >
            <option value="__all__">Todas</option>
            {brandOptions.map((b) => (
              <option key={b} value={b}>
                {b} ({facetBrands[b] || 0})
              </option>
            ))}
          </select>
        </div>
        <div>
          <label className="block text-xs text-slate-600 mb-1">Categoría</label>
          <select
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 shadow-sm"
            value={typeFilter || "__all__"}
            onChange={(e) =>
              setTypeFilter(e.target.value === "__all__" ? null : e.target.value)
            }
          >
            <option value="__all__">Todas</option>
            {typeOptions.map((t) => (
              <option key={t} value={t}>
                {t} ({facetTypes[t] || 0})
              </option>
            ))}
          </select>
        </div>
        <div className="col-span-2">
          <label className="block text-xs text-slate-600 mb-1">Ordenar</label>
          <select
            className="w-full rounded-xl border border-slate-200 bg-white px-3 py-2 shadow-sm"
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value)}
          >
            <option value="relevance">Relevancia</option>
            <option value="az">Alfabético A–Z</option>
            <option value="priceAsc">Precio ↑</option>
            <option value="priceDesc">Precio ↓</option>
            <option value="popular">Más usados</option>
          </select>
        </div>
      </div>

      <section>
        {sorted.map((p) => {
          const isFav = favorites.has(p.sku);
          return (
            <article
              key={p.sku}
              className="mx-4 mb-4 rounded-2xl border border-slate-200 bg-white shadow-sm overflow-hidden"
            >
              <div className="relative">
                <button
                  className="block w-full"
                  onClick={() => onOpenDetail(p)}
                  aria-label="Ver detalle"
                >
                  <img
                    src={isHttpUrl(p.image) ? p.image : FALLBACK_IMAGE}
                    alt={p.name}
                    className="w-full h-40 object-cover bg-slate-50"
                  />
                </button>
                <div className="absolute left-2 top-2 flex flex-wrap gap-2">
                  {(p.badges || []).map((b) => (
                    <span
                      key={b}
                      className={`text-[10px] px-2 py-0.5 rounded-full border ${badgeClass(
                        b
                      )}`}
                    >
                      {b}
                    </span>
                  ))}
                  {typeof p.oldPrice === "number" &&
                    (p.price || 0) > 0 &&
                    p.oldPrice > (p.price || 0) && (
                      <span
                        className={`text-[10px] px-2 py-0.5 rounded-full border ${badgeClass(
                          "Oferta"
                        )}`}
                      >
                        {Math.round(((p.oldPrice - p.price) / p.oldPrice) * 100)}
                        % OFF
                      </span>
                    )}
                </div>
                <button
                  className="absolute right-2 top-2 w-8 h-8 rounded-full bg-white/90 shadow flex items-center justify-center text-sm"
                  onClick={(e) => {
                    e.stopPropagation();
                    onToggleFavorite(p.sku);
                  }}
                  aria-label={isFav ? "Quitar de favoritos" : "Agregar a favoritos"}
                >
                  {isFav ? "★" : "☆"}
                </button>
              </div>
              <div className="p-3">
                <div className="font-medium leading-snug">{p.name}</div>
                <div className="text-xs text-slate-500 mb-2">
                  {p.brand || "(sin marca)"} · {p.type || "(sin categoría)"}
                </div>
                <div className="flex items-baseline gap-2 mb-2">
                  <div className="font-semibold">
                    ${(p.price || 0).toLocaleString("es-CL")}
                  </div>
                  {typeof p.oldPrice === "number" &&
                    (p.price || 0) > 0 &&
                    p.oldPrice > (p.price || 0) && (
                      <div className="text-xs text-slate-400 line-through">
                        ${p.oldPrice.toLocaleString("es-CL")}
                      </div>
                    )}
                </div>
                <div className="text-[11px] text-slate-500 flex flex-wrap gap-1 mb-3">
                  {(p.tags || []).slice(0, 4).map((t) => (
                    <span key={t}>#{t}</span>
                  ))}
                </div>
                <div className="flex justify-between items-center">
                  <button
                    onClick={() => onOpenDetail(p)}
                    className="rounded-xl border border-slate-200 bg-white text-sm px-3 py-2 hover:bg-slate-50"
                  >
                    ℹ️ Detalle
                  </button>
                  <button
                    onClick={() => onAddToCart(p)}
                    className="rounded-xl bg-slate-900 text-white text-sm px-3 py-2 hover:bg-slate-800 active:scale-[0.99]"
                  >
                    Agregar al carrito
                  </button>
                </div>
              </div>
            </article>
          );
        })}
        {sorted.length === 0 && (
          <div className="text-center text-slate-500 text-sm py-12">
            Sin resultados con los filtros actuales.
          </div>
        )}
      </section>
    </>
  );
}

interface DebugViewProps {
  previewHeaders: string[];
  previewData: NormalizedRow[];
  ignoreShowFlag: boolean;
  onImportFromCache: () => void;
  onClose: () => void;
}

function DebugView({
  previewHeaders,
  previewData,
  ignoreShowFlag,
  onImportFromCache,
  onClose,
}: DebugViewProps) {
  return (
    <div className="px-4 pb-24">
      <h2 className="text-sm font-semibold mb-2">Vista Previa (cache)</h2>
      <div className="text-xs text-slate-600 mb-2">
        <b>Cabeceras:</b> {previewHeaders.join(" · ") || "(sin)"} —{" "}
        <b>Filas:</b> {previewData.length}
      </div>

      <div className="rounded-xl border border-slate-200 bg-white p-3 mb-4 overflow-auto">
        <div className="text-xs font-medium mb-2">Grilla (máx 5 filas)</div>
        {previewHeaders.length > 0 && (
          <table className="min-w-full text-[11px] border">
            <thead>
              <tr>
                {previewHeaders.map((h) => (
                  <th
                    key={h}
                    className="border px-2 py-1 bg-slate-50 text-left"
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {previewData.slice(0, 5).map((row, ri) => (
                <tr key={ri}>
                  {previewHeaders.map((h) => (
                    <td key={h} className="border px-2 py-1">
                      {String((row as any)?.[h] ?? "")}
                    </td>
                  ))}
                </tr>
              ))}
              {previewData.length === 0 && (
                <tr>
                  <td
                    className="px-2 py-2 text-slate-500"
                    colSpan={previewHeaders.length || 1}
                  >
                    Sin filas
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        )}
        {previewHeaders.length === 0 && (
          <div className="text-xs text-slate-500">Sin cabeceras detectadas.</div>
        )}
        <div className="mt-2 flex gap-2">
          <button
            className="px-3 py-2 rounded-lg bg-slate-900 text-white text-sm"
            onClick={onImportFromCache}
          >
            Importar cache y volver
          </button>
          <button
            className="px-3 py-2 rounded-lg border text-sm"
            onClick={onClose}
          >
            Volver
          </button>
        </div>
      </div>

      <div className="rounded-xl border border-slate-200 bg-white p-3">
        <div className="text-xs font-medium mb-2">
          normalizeIncoming() — 10 primeras
        </div>
        {previewData.slice(0, 10).map((row, idx) => {
          const normObj = normalizeIncoming(row, { ignoreShowFlag });
          return (
            <div key={idx} className="mb-2">
              <div className="text-[11px] text-slate-500 mb-1">
                Row #{idx + 1}
              </div>
              <div className="text-[11px] rounded-lg bg-slate-50 border border-slate-200 p-2 overflow-auto mb-1">
                RAW: {JSON.stringify(row)}
              </div>
              <div
                className={
                  "text-[11px] rounded-lg p-2 overflow-auto " +
                  (normObj
                    ? "bg-emerald-50 border border-emerald-200"
                    : "bg-rose-50 border-rose-200")
                }
              >
                {normObj ? (
                  <>Normalized: {JSON.stringify(normObj)}</>
                ) : (
                  <>
                    Descartada (flag/valores):{" "}
                    {JSON.stringify({
                      name: (row as any)?.Producto || (row as any)?.producto,
                      mostrar: (row as any)?.["Mostrar en Tienda"],
                    })}
                  </>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

interface FavoritesViewProps {
  data: Product[];
  favorites: Set<string>;
  onAddToCart: (p: Product) => void;
}

function FavoritesView({ data, favorites, onAddToCart }: FavoritesViewProps) {
  const favItems = data.filter((p) => favorites.has(p.sku));

  return (
    <div className="px-4 pb-6">
      <h2 className="text-sm text-slate-600 mb-3">
        Favoritos ({favItems.length})
      </h2>
      {favItems.length === 0 && (
        <div className="text-sm text-slate-500">No tienes favoritos aún.</div>
      )}
      {favItems.map((p) => (
        <div
          key={p.sku}
          className="mb-3 rounded-xl border border-slate-200 bg-white p-3 flex items-center justify-between"
        >
          <div className="flex items-center gap-3">
            <img
              src={isHttpUrl(p.image) ? p.image : FALLBACK_IMAGE}
              alt={p.name}
              className="w-14 h-14 object-cover rounded-lg"
            />
            <div>
              <div className="text-sm font-medium leading-snug line-clamp-2 max-w-[220px]">
                {p.name}
              </div>
              <div className="text-[11px] text-slate-500">
                {p.brand || "(sin marca)"} · {p.type || "(sin categoría)"}
              </div>
            </div>
          </div>
          <button
            className="text-sm px-3 py-1 rounded-lg bg-slate-900 text-white"
            onClick={() => onAddToCart(p)}
          >
            Agregar
          </button>
        </div>
      ))}
    </div>
  );
}

interface CartViewProps {
  data: Product[];
  cart: CartItem[];
  setCart: React.Dispatch<React.SetStateAction<CartItem[]>>;
}

function CartView({ data, cart, setCart }: CartViewProps) {
  const lines = cart
    .map((l) => ({
      ...l,
      prod: data.find((p) => p.sku === l.sku),
    }))
    .filter((l) => l.prod) as Array<CartItem & { prod: Product }>;

  const subtotal = lines.reduce(
    (s, l) => s + ((l.prod.price || 0) * l.qty),
    0
  );

  return (
    <div className="px-4 pb-6">
      <h2 className="text-sm text-slate-600 mb-3">
        Carrito ({lines.length} ítem(s))
      </h2>
      {lines.length === 0 && (
        <div className="text-sm text-slate-500">Tu carrito está vacío.</div>
      )}
      {lines.map((l) => (
        <div
          key={l.sku}
          className="mb-3 rounded-xl border border-slate-200 bg-white p-3 flex items-center justify-between"
        >
          <div className="flex items-center gap-3">
            <img
              src={isHttpUrl(l.prod.image) ? l.prod.image : FALLBACK_IMAGE}
              alt={l.prod.name}
              className="w-14 h-14 object-cover rounded-lg"
            />
            <div>
              <div className="text-sm font-medium leading-snug line-clamp-2 max-w-[200px]">
                {l.prod.name}
              </div>
              <div className="text-[11px] text-slate-500">
                {l.prod.brand || "(sin marca)"} ·{" "}
                {l.prod.type || "(sin categoría)"}
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              className="px-2 py-1 rounded-lg border"
              onClick={() =>
                setCart((cur) => {
                  const idx = cur.findIndex((x) => x.sku === l.sku);
                  if (idx < 0) return cur;
                  const next = [...cur];
                  const newQty = next[idx].qty - 1;
                  if (newQty <= 0) next.splice(idx, 1);
                  else next[idx] = { ...next[idx], qty: newQty };
                  return next;
                })
              }
            >
              -
            </button>
            <div className="w-6 text-center text-sm">{l.qty}</div>
            <button
              className="px-2 py-1 rounded-lg border"
              onClick={() =>
                setCart((cur) => {
                  const idx = cur.findIndex((x) => x.sku === l.sku);
                  if (idx < 0) return [...cur, { sku: l.sku, qty: 1 }];
                  const next = [...cur];
                  next[idx] = { ...next[idx], qty: next[idx].qty + 1 };
                  return next;
                })
              }
            >
              +
            </button>
          </div>
        </div>
      ))}
      {lines.length > 0 && (
        <div className="mt-4 rounded-xl border border-slate-200 bg-white p-3 flex items-center justify-between">
          <div className="text-sm">Subtotal</div>
          <div className="font-semibold">
            ${subtotal.toLocaleString("es-CL")}
          </div>
        </div>
      )}
    </div>
  );
}

// ===== Componente principal =====
export default function App() {
  const [data, setData] = useState<Product[]>(DEFAULT_PRODUCTS);
  const [queryInput, setQueryInput] = useState<string>("");
  const qDeferred = useDeferredValue<string>(queryInput);
  const [brandFilter, setBrandFilter] = useState<string | null>(null);
  const [typeFilter, setTypeFilter] = useState<string | null>(null);
  const [quickSel] = useState<Set<string>>(new Set()); // reservado
  const [favorites, setFavorites] = useState<Set<string>>(new Set());
  const [cart, setCart] = useState<CartItem[]>([]);
  const [toast, setToast] = useState<{ show: boolean; text: string }>({
    show: false,
    text: "",
  });
  const [importOpen, setImportOpen] = useState<boolean>(false);
  const [importMode, setImportMode] = useState<"append" | "replace">("append");
  const [jsonText, setJsonText] = useState<string>("");
  const [tab, setTab] = useState<"catalog" | "fav" | "cart">("catalog");
  const [sortBy, setSortBy] = useState<string>("relevance");
  const [popMap, setPopMap] = useState<Record<string, number>>({});
  const [detailOpen, setDetailOpen] = useState<boolean>(false);
  const [detailProd, setDetailProd] = useState<Product | null>(null);
  const cartBtnRef = useRef<HTMLButtonElement | null>(null);

  // === Voz: STT + TTS ===
  const [hasSpeechApi, setHasSpeechApi] = useState<boolean>(false);
  const [listening, setListening] = useState<boolean>(false);
  const [lastTranscript, setLastTranscript] = useState<string>("");
  const [ttsEnabled, setTtsEnabled] = useState<boolean>(true);
  const recognitionRef = useRef<any | null>(null);

  function speak(text: string) {
    if (!ttsEnabled) return;
    if (typeof window === "undefined") return;
    const synth = window.speechSynthesis;
    if (!synth) return;
    const utter = new SpeechSynthesisUtterance(text);
    utter.lang = "es-ES";
    synth.speak(utter);
  }

  function handleVoiceCommand(raw: string) {
    const text = raw.toLowerCase().trim();
    if (!text) return;

    // Navegación básica
    if (text.includes("carrito")) {
      setTab("cart");
      speak("Abro tu carrito.");
      showToast("Te muestro el carrito");
      return;
    }
    if (text.includes("favorito") || text.includes("favoritos")) {
      setTab("fav");
      speak("Voy a tus favoritos.");
      showToast("Te muestro tus favoritos");
      return;
    }
    if (
      text.includes("catálogo") ||
      text.includes("catalogo") ||
      text.includes("productos") ||
      text.includes("inicio")
    ) {
      setTab("catalog");
      speak("Vuelvo al catálogo.");
      showToast("Catálogo de productos");
      return;
    }

    // Limpiar búsqueda
    if (
      text.includes("limpiar búsqueda") ||
      text.includes("limpia búsqueda") ||
      text.includes("borrar búsqueda") ||
      text.includes("limpiar busqueda") ||
      text.includes("borrar filtros") ||
      text.includes("limpiar filtros")
    ) {
      setQueryInput("");
      speak("Limpio la búsqueda.");
      showToast("Búsqueda y filtros limpiados (búsqueda).");
      return;
    }

    // Buscar productos
    if (text.startsWith("buscar ")) {
      const q = text.replace(/^buscar\s+/, "");
      setTab("catalog");
      setQueryInput(q);
      speak(`Busco productos que coincidan con ${q}.`);
      showToast(`Buscando: ${q}`);
      return;
    }
    if (text.startsWith("busca ")) {
      const q = text.replace(/^busca\s+/, "");
      setTab("catalog");
      setQueryInput(q);
      speak(`Busco productos que coincidan con ${q}.`);
      showToast(`Buscando: ${q}`);
      return;
    }

    // Agregar al carrito por nombre
    if (
      text.includes("agrega ") ||
      text.includes("añade ") ||
      text.includes("agregar ")
    ) {
      const triggerIndex = Math.min(
        ...["agrega ", "añade ", "agregar "]
          .map((w) => text.indexOf(w))
          .filter((i) => i >= 0)
      );
      if (Number.isFinite(triggerIndex)) {
        const fragment = text
          .slice(triggerIndex)
          .replace("agrega", "")
          .replace("añade", "")
          .replace("agregar", "")
          .replace("al carrito", "")
          .replace("a favoritos", "")
          .trim();

        if (fragment.length > 1) {
          const match = data.find((p) =>
            p.name.toLowerCase().includes(fragment)
          );
          if (match) {
            addToCart(match);
            speak(`Agrego ${match.name} al carrito.`);
            return;
          }
        }
      }
    }

    // Estado de pedido / reclamo (demo)
    if (
      (text.includes("estado") || text.includes("seguimiento")) &&
      (text.includes("pedido") || text.includes("orden"))
    ) {
      const msg =
        "Por ahora soy una demo y no tengo acceso a pedidos reales. Puedo ayudarte con los productos de tu carrito.";
      speak(msg);
      showToast("Demo: sin conexión a pedidos reales todavía.");
      return;
    }

    if (text.includes("reclamo") || text.includes("queja")) {
      const msg =
        "Puedo ayudarte a identificar el producto y el problema, pero esta demo aún no envía reclamos a un sistema real.";
      speak(msg);
      showToast("Demo reclamos: aquí se podría abrir un formulario o chat.");
      return;
    }

    // Fallback: usar como búsqueda
    setTab("catalog");
    setQueryInput(text);
    speak(`Uso eso como búsqueda: ${text}.`);
    showToast(`Buscando: ${text}`);
  }

  useEffect(() => {
    if (typeof window === "undefined") return;
    const SR =
      (window as any).SpeechRecognition ||
      (window as any).webkitSpeechRecognition;
    if (!SR) {
      setHasSpeechApi(false);
      return;
    }
    setHasSpeechApi(true);
    const rec = new SR();
    rec.lang = "es-ES";
    rec.interimResults = false;
    rec.maxAlternatives = 1;

    rec.onresult = (event: any) => {
      try {
        const t = event.results[0][0].transcript as string;
        setLastTranscript(t);
        handleVoiceCommand(t);
      } catch {
        // noop
      }
    };
    rec.onerror = () => {
      setListening(false);
    };
    rec.onend = () => {
      setListening(false);
    };

    recognitionRef.current = rec;

    return () => {
      try {
        rec.stop();
      } catch {
        // noop
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function toggleListening() {
    if (!hasSpeechApi) {
      showToast("Tu navegador no soporta dictado por voz.");
      return;
    }
    const rec = recognitionRef.current;
    if (!rec) {
      showToast("La API de voz no está lista.");
      return;
    }
    if (listening) {
      try {
        rec.stop();
      } catch {
        // noop
      }
      setListening(false);
    } else {
      try {
        setListening(true);
        rec.start();
        showToast("Escuchando… habla cerca del micrófono.");
      } catch {
        setListening(false);
      }
    }
  }

  // Debug / Preview cache
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [previewHeaders, setPreviewHeaders] = useState<string[]>([]);
  const [previewRows, setPreviewRows] = useState<NormalizedRow[]>([]);
  const [previewData, setPreviewData] = useState<NormalizedRow[]>([]);
  const [ignoreShowFlag] = useState<boolean>(false);
  const [view, setView] = useState<"main" | "debug">("main");

  function showToast(text: string, ms = 1600) {
    setToast({ show: true, text });
    setTimeout(() => setToast({ show: false, text: "" }), ms);
  }

  function toggleFavorite(sku: string) {
    setFavorites((prev) => {
      const next = new Set(prev);
      if (next.has(sku)) next.delete(sku);
      else next.add(sku);
      return next;
    });
  }

  function addToCart(product: Product) {
    setCart((cur) => {
      const idx = cur.findIndex((l) => l.sku === product.sku);
      if (idx >= 0) {
        const next = [...cur];
        next[idx] = { ...next[idx], qty: next[idx].qty + 1 };
        return next;
      }
      return [...cur, { sku: product.sku, qty: 1 }];
    });
    setPopMap((m) => ({
      ...m,
      [product.sku]: (m[product.sku] || 0) + 1,
    }));
    try {
      cartBtnRef.current?.classList.add("ring-2", "ring-rose-300");
      setTimeout(() => {
        cartBtnRef.current?.classList.remove("ring-2", "ring-rose-300");
      }, 600);
    } catch {
      // noop
    }
    showToast("Agregado al carrito");
  }

  // Filtro base
  const baseFiltered = useMemo(() => {
    const q = String(qDeferred || "").trim().toLowerCase();
    return data.filter((p) => {
      const matchesQ =
        !q ||
        p.name.toLowerCase().includes(q) ||
        (p.brand || "").toLowerCase().includes(q) ||
        (p.type || "").toLowerCase().includes(q) ||
        (p.tags || []).some((t) => t.toLowerCase().includes(q));
      const matchesBrand = !brandFilter || p.brand === brandFilter;
      const matchesType = !typeFilter || p.type === typeFilter;
      const matchesQuick =
        quickSel.size === 0 ||
        Array.from(quickSel).every(
          (f) => (p.tags || []).includes(f) || p.type === f
        );
      return matchesQ && matchesBrand && matchesType && matchesQuick;
    });
  }, [data, qDeferred, brandFilter, typeFilter, quickSel]);

  const facetBrands = useMemo(
    () => countBy(baseFiltered.map((p) => (p.brand || "(sin marca)").trim())),
    [baseFiltered]
  );
  const facetTypes = useMemo(
    () => countBy(baseFiltered.map((p) => (p.type || "(sin categoría)").trim())),
    [baseFiltered]
  );

  const sorted = useMemo(() => {
    const arr = [...baseFiltered];
    switch (sortBy) {
      case "az":
        arr.sort((a, b) => a.name.localeCompare(b.name, "es"));
        break;
      case "priceAsc":
        arr.sort((a, b) => (a.price || 0) - (b.price || 0));
        break;
      case "priceDesc":
        arr.sort((a, b) => (b.price || 0) - (a.price || 0));
        break;
      case "popular":
        arr.sort(
          (a, b) => (popMap[b.sku] || 0) - (popMap[a.sku] || 0)
        );
        break;
      default:
        break;
    }
    return arr;
  }, [baseFiltered, sortBy, popMap]);

  // Importación / Vista previa
  function pickBestSheet(wb: XLSX.WorkBook): string {
    let best = wb.SheetNames[0];
    let bestScore = -1;
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const rows = (XLSX.utils.sheet_to_json(ws, {
        header: 1,
        defval: "",
      }) as unknown[][]) || [];
      const nonEmpty = rows.filter(
        (r) =>
          Array.isArray(r) &&
          r.some((c) => String(c ?? "").trim() !== "")
      ).length;
      const headerPresence = rows
        .slice(0, 5)
        .some((r) =>
          (r || []).some((c) => {
            const n = normalizeKey(c);
            return [
              "sku",
              "producto",
              "marca",
              "categoria",
              "categoría",
              "precio base",
              "url imagen",
              "url_imagen",
            ].includes(n);
          })
        )
        ? 10
        : 0;
      const score = nonEmpty + headerPresence;
      if (score > bestScore) {
        bestScore = score;
        best = name;
      }
    }
    return best;
  }

  async function previewFile(file: File) {
    if (!file) {
      showToast("Selecciona un archivo primero");
      return;
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const best = pickBestSheet(wb);
    const ws = wb.Sheets[best];

    let json = XLSX.utils.sheet_to_json(ws, { defval: "" }) as NormalizedRow[];
    let usedDetector = false;

    if (!Array.isArray(json) || json.length === 0) {
      usedDetector = true;
      const { headers, rows, reason } = detectTableWithData(ws);
      if (headers.length) {
        setPreviewHeaders(headers);
        setPreviewRows(rows.slice(0, 3));
        setPreviewData(rows);
        showToast(`Vista previa: hoja "${best}", ${rows.length} fila(s)`, 2400);
        if (rows.length === 0 && reason) showToast(reason, 3200);
        return;
      } else {
        const rows2 =
          (XLSX.utils.sheet_to_json(ws, {
            header: 1,
            defval: "",
          }) as unknown[][]) || [];
        if (Array.isArray(rows2) && rows2.length > 1) {
          const headers = rows2[0].map((h) => String(h || "").trim());
          const rows = rows2
            .slice(1)
            .map((r) => {
              const obj: NormalizedRow = {};
              headers.forEach((h, i) => {
                (obj as any)[h || `col_${i + 1}`] = (r as unknown[])[i];
              });
              return obj;
            })
            .filter((o) =>
              Object.values(o).some((v) => String(v ?? "").trim() !== "")
            );
          setPreviewHeaders(headers);
          setPreviewRows(rows.slice(0, 3));
          setPreviewData(rows);
          showToast(`Vista previa: hoja "${best}", ${rows.length} fila(s)`, 2400);
          return;
        }
        setPreviewHeaders([]);
        setPreviewRows([]);
        setPreviewData([]);
        showToast(
          `No se pudo detectar cabeceras/filas${reason ? `: ${reason}` : ""}`,
          3200
        );
        return;
      }
    }

    const rows = json.filter((o) =>
      Object.values(o).some((v) => String(v ?? "").trim() !== "")
    );
    const headers = rows[0] ? Object.keys(rows[0]) : [];
    setPreviewHeaders(headers);
    setPreviewRows(rows.slice(0, 3));
    setPreviewData(rows);
    showToast(
      `Vista previa: hoja "${best}", ${rows.length} fila(s)${
        usedDetector ? " (detector)" : " (directo)"
      }`,
      2400
    );
  }

  function importAndMerge(list: NormalizedRow[]) {
    const items: Product[] = [];
    for (const r of list || []) {
      const p = normalizeIncoming(r, { ignoreShowFlag });
      if (p) items.push(p);
    }
    if (items.length === 0) {
      const sample =
        list && list[0] && typeof list[0] === "object"
          ? Object.keys(list[0]).slice(0, 12).join(", ")
          : "(sin encabezados)";
      showToast(
        `No se reconocieron filas válidas. Encabezados detectados: ${sample}`,
        4200
      );
      return;
    }
    if (importMode === "replace") {
      setBrandFilter(null);
      setTypeFilter(null);
      setQueryInput("");
      setData(items);
    } else {
      setData((cur) => [...cur, ...items]);
    }
    showToast(`Importados ${items.length} producto(s)`);
  }

  function handleJsonFile(file: File, forPreview = false) {
    const fr = new FileReader();
    fr.onload = () => {
      try {
        const arr = JSON.parse(String(fr.result || ""));
        const rows = Array.isArray(arr) ? (arr as NormalizedRow[]) : [];
        if (forPreview) {
          const headers = rows[0] ? Object.keys(rows[0]) : [];
          setPreviewHeaders(headers);
          setPreviewRows(rows.slice(0, 3));
          setPreviewData(rows);
          showToast(`Vista previa JSON: ${rows.length} fila(s)`);
        } else {
          if (previewData.length > 0) {
            importAndMerge(previewData);
            setImportOpen(false);
            return;
          }
          importAndMerge(rows);
          setImportOpen(false);
        }
      } catch {
        showToast("JSON inválido");
      }
    };
    fr.readAsText(file);
  }

  function handleFile(forPreview = false) {
    const file = selectedFile;
    if (!file) {
      showToast("Selecciona un archivo");
      return;
    }
    const isJson = /json$/i.test(file.type) || /\.json$/i.test(file.name);
    if (isJson) return handleJsonFile(file, forPreview);
    if (forPreview) return previewFile(file);

    if (previewData.length > 0) {
      importAndMerge(previewData);
      setImportOpen(false);
      return;
    }

    const fr = new FileReader();
    fr.onload = (e) => {
      try {
        const buf = (e.target as FileReader).result as ArrayBuffer;
        const wb = XLSX.read(buf, { type: "array" });
        const best = pickBestSheet(wb);
        const ws = wb.Sheets[best];

        let json = XLSX.utils.sheet_to_json(ws, { defval: "" }) as NormalizedRow[];
        if (!Array.isArray(json) || json.length === 0) {
          const { headers, rows } = detectTableWithData(ws);
          if (headers.length) {
            importAndMerge(rows);
            setImportOpen(false);
            return;
          }
          const rows2 =
            (XLSX.utils.sheet_to_json(ws, {
              header: 1,
              defval: "",
            }) as unknown[][]) || [];
          if (Array.isArray(rows2) && rows2.length > 1) {
            const headers2 = rows2[0].map((h) => String(h || "").trim());
            const rows = rows2
              .slice(1)
              .map((r) => {
                const obj: NormalizedRow = {};
                headers2.forEach((h, i) => {
                  (obj as any)[h || `col_${i + 1}`] = (r as unknown[])[i];
                });
                return obj;
              })
              .filter((o) =>
                Object.values(o).some(
                  (v) => String(v ?? "").trim() !== ""
                )
              );
            importAndMerge(rows);
            setImportOpen(false);
            return;
          }
          showToast("Archivo no válido o sin filas");
          return;
        }
        importAndMerge(json);
        setImportOpen(false);
      } catch {
        showToast("Archivo no válido (XLSX/CSV/JSON)");
      }
    };
    fr.readAsArrayBuffer(file);
  }

  function applyImportFromTextarea() {
    try {
      const arr = JSON.parse(jsonText);
      const rows = Array.isArray(arr) ? (arr as NormalizedRow[]) : [];
      setPreviewHeaders(rows[0] ? Object.keys(rows[0]) : []);
      setPreviewRows(rows.slice(0, 3));
      setPreviewData(rows);
      importAndMerge(rows);
      setImportOpen(false);
    } catch {
      try {
        const wb = XLSX.read(jsonText, { type: "string" });
        const best = pickBestSheet(wb);
        const ws = wb.Sheets[best];
        const { headers, rows, reason } = detectTableWithData(ws);
        if (headers.length) {
          setPreviewHeaders(headers);
          setPreviewRows(rows.slice(0, 3));
          setPreviewData(rows);
          importAndMerge(rows);
          setImportOpen(false);
        } else {
          showToast(
            `Contenido pegado inválido${reason ? `: ${reason}` : ""}`
          );
        }
      } catch {
        showToast("Contenido pegado inválido (JSON o CSV/TSV)");
      }
    }
  }

  const totalItemsCart = cart.reduce((s, i) => s + i.qty, 0);

  return (
    <div className="min-h-screen w-full flex justify-center bg-gradient-to-b from-rose-50 via-white to-slate-50">
      <div className="w-full max-w-[430px] min-h-screen relative pb-36">
        {toast.show && (
          <div className="fixed top-4 left-1/2 -translate-x-1/2 z-[60] bg-slate-900 text-white text-xs px-3 py-2 rounded-full shadow-lg">
            {toast.text}
          </div>
        )}

        <header className="px-4 pt-4 pb-2 flex items-center gap-3 sticky top-0 z-[40] bg-gradient-to-b from-rose-50/90 via-white/90 to-white/80 backdrop-blur">
          <div className="w-10 h-10 rounded-2xl bg-red-600 text-white grid place-items-center font-bold shadow">
            N
          </div>
          <div className="font-semibold flex-1">
            Nestlé · Despensa inteligente
          </div>

          <button
            className={
              "text-xs px-2 py-1 rounded-lg border mr-2 " +
              (view === "debug"
                ? "bg-slate-900 text-white"
                : "bg-white")
            }
            onClick={() => setView(view === "debug" ? "main" : "debug")}
          >
            Debug
          </button>

          <button
            className={
              "text-xs px-2 py-1 rounded-lg border mr-2 " +
              (ttsEnabled
                ? "bg-emerald-50 border-emerald-300 text-emerald-800"
                : "bg-white text-slate-600")
            }
            onClick={() => setTtsEnabled((v) => !v)}
            title={
              ttsEnabled
                ? "Desactivar respuestas habladas"
                : "Activar respuestas habladas"
            }
          >
            🔊
          </button>

          <button
            className="text-xs px-2 py-1 rounded-lg border border-slate-200 bg-white shadow-sm hover:bg-slate-50"
            onClick={() => setImportOpen(true)}
            title="Importar Excel/CSV/JSON"
          >
            Datos
          </button>
          <div className="ml-2 text-slate-500 text-sm">
            [{baseFiltered.length}]
          </div>
        </header>

        {view === "debug" ? (
          <DebugView
            previewHeaders={previewHeaders}
            previewData={previewData}
            ignoreShowFlag={ignoreShowFlag}
            onImportFromCache={() => {
              if (previewData.length > 0) {
                importAndMerge(previewData);
                setView("main");
              } else {
                showToast("No hay datos en caché");
              }
            }}
            onClose={() => setView("main")}
          />
        ) : tab === "catalog" ? (
          <CatalogView
            facetBrands={facetBrands}
            facetTypes={facetTypes}
            sorted={sorted}
            queryInput={queryInput}
            setQueryInput={setQueryInput}
            brandFilter={brandFilter}
            setBrandFilter={setBrandFilter}
            typeFilter={typeFilter}
            setTypeFilter={setTypeFilter}
            sortBy={sortBy}
            setSortBy={setSortBy}
            favorites={favorites}
            onToggleFavorite={toggleFavorite}
            onOpenDetail={(p) => {
              setDetailProd(p);
              setDetailOpen(true);
            }}
            onAddToCart={addToCart}
          />
        ) : tab === "fav" ? (
          <FavoritesView
            data={data}
            favorites={favorites}
            onAddToCart={addToCart}
          />
        ) : (
          <CartView data={data} cart={cart} setCart={setCart} />
        )}

        <footer className="fixed bottom-0 left-0 right-0 grid place-items-center z-[50] pointer-events-none">
          <div className="w-full max-w-[430px] px-4 pb-3 pointer-events-auto">
            <nav className="rounded-2xl border border-slate-200 bg-white shadow-md px-3 py-2 flex items-center justify-between text-sm">
              <button
                className={
                  "px-3 py-2 rounded-xl " +
                  (tab === "catalog"
                    ? "bg-slate-900 text-white"
                    : "hover:bg-slate-100")
                }
                onClick={() => setTab("catalog")}
              >
                🏪 Catálogo
              </button>
              <button
                className={
                  "px-3 py-2 rounded-xl " +
                  (tab === "fav"
                    ? "bg-slate-900 text-white"
                    : "hover:bg-slate-100")
                }
                onClick={() => setTab("fav")}
              >
                ⭐ Favoritos ({favorites.size})
              </button>
              <button
                ref={cartBtnRef}
                className={
                  "px-3 py-2 rounded-xl " +
                  (tab === "cart"
                    ? "bg-slate-900 text-white"
                    : "hover:bg-slate-100")
                }
                onClick={() => setTab("cart")}
              >
                🧺 Carrito ({totalItemsCart})
              </button>
            </nav>
          </div>
        </footer>

        {importOpen && (
          <div className="fixed inset-0 bg-black/30 backdrop-blur-sm z-[70] grid place-items-center p-4">
            <div className="w-full max-w-[720px] bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
              <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
                <div className="font-medium">
                  Importar datos (Excel / CSV / JSON / Pegar)
                </div>
                <button
                  className="text-slate-500"
                  onClick={() => setImportOpen(false)}
                >
                  ✕
                </button>
              </div>
              <div className="p-4 space-y-4">
                <div
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    const files = e.dataTransfer.files;
                    const f = files && files[0];
                    if (f) {
                      setSelectedFile(f);
                      showToast(`Archivo: ${f.name}`);
                    }
                  }}
                  className="rounded-2xl border-2 border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-center"
                >
                  <div className="text-sm font-medium mb-1">
                    1) Selecciona o arrastra tu archivo
                  </div>
                  <div className="text-xs text-slate-500">
                    Soporta .xlsx, .csv y .json — también puedes pegar abajo
                  </div>
                  <div className="mt-3 flex items-center justify-center gap-3 text-xs">
                    <input
                      id="file-input-hidden"
                      type="file"
                      className="hidden"
                      accept=".xlsx,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/json,text/csv"
                      onChange={(e) => {
                        const f =
                          e.target.files && e.target.files[0];
                        if (f) {
                          setSelectedFile(f);
                          showToast(`Archivo: ${f.name}`);
                        }
                      }}
                    />
                    <button
                      className="px-3 py-2 rounded-lg border bg-white"
                      onClick={() => {
                        const el = document.getElementById(
                          "file-input-hidden"
                        ) as HTMLInputElement | null;
                        el?.click();
                      }}
                    >
                      Elegir archivo…
                    </button>
                    <div className="ml-3 flex items-center text-xs">
                      <label className="flex items-center gap-1">
                        <input
                          type="radio"
                          name="mode"
                          checked={importMode === "append"}
                          onChange={() => setImportMode("append")}
                        />
                        Append
                      </label>
                      <label className="flex items-center gap-1 ml-3">
                        <input
                          type="radio"
                          name="mode"
                          checked={importMode === "replace"}
                          onChange={() => setImportMode("replace")}
                        />
                        Replace
                      </label>
                    </div>
                  </div>
                  <div className="mt-3 flex items-center justify-center gap-2">
                    <button
                      className="px-3 py-2 rounded-lg bg-slate-900 text-white"
                      onClick={() => selectedFile && handleFile(true)}
                    >
                      Vista previa
                    </button>
                    <button
                      className="px-3 py-2 rounded-lg border bg-white"
                      onClick={() => selectedFile && handleFile(false)}
                    >
                      Importar
                    </button>
                  </div>
                  {selectedFile && (
                    <div className="mt-2 text-xs text-slate-600">
                      Seleccionado: <b>{selectedFile.name}</b>
                    </div>
                  )}
                </div>

                <div className="rounded-xl border border-slate-200 bg-white p-3">
                  <div className="text-xs text-slate-600 mb-1">
                    2) (Opcional) Pegar datos JSON o CSV/TSV de Excel
                  </div>
                  <textarea
                    className="w-full h-28 rounded-xl border border-slate-200 bg-white px-3 py-2 text-xs font-mono"
                    value={jsonText}
                    onChange={(e) => setJsonText(e.target.value)}
                    placeholder='[{"SKU":"123","Producto":"...","Marca":"...","Precio Base":1990,"URL_Imagen":"..."}]'
                  />
                  <div className="flex justify-end gap-2 pt-1">
                    <button
                      className="px-3 py-2 text-sm rounded-lg border border-slate-200 bg-white"
                      onClick={() => setImportOpen(false)}
                    >
                      Cerrar
                    </button>
                    <button
                      className="px-3 py-2 text-sm rounded-lg bg-slate-900 text-white"
                      onClick={applyImportFromTextarea}
                    >
                      Importar desde pegado
                    </button>
                  </div>
                </div>

                {(previewHeaders.length > 0 || previewRows.length > 0) && (
                  <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-3">
                    <div className="text-sm font-medium mb-2">
                      Vista previa detectada
                    </div>
                    <div className="text-xs mb-2">
                      <b>Cabeceras</b>:{" "}
                      {previewHeaders.join(" · ") || "(ninguna)"} ·{" "}
                      <b>Filas</b>: {previewData.length}
                    </div>
                    <div className="text-xs">
                      <div className="font-medium mb-1">
                        Primeras filas (máx 3):
                      </div>
                      <div className="grid gap-1">
                        {previewRows.map((r, idx) => (
                          <div
                            key={idx}
                            className="rounded-lg bg-white border border-emerald-200 p-2 overflow-auto text-[11px]"
                          >
                            {JSON.stringify(r)}
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="mt-2">
                      <button
                        className="px-3 py-2 rounded-lg border bg-white text-sm"
                        onClick={() => {
                          setView("debug");
                          setImportOpen(false);
                        }}
                      >
                        Abrir Debug
                      </button>
                    </div>
                  </div>
                )}

                <div className="text-[11px] text-slate-500">
                  <div className="font-medium mb-1">
                    Columnas reconocidas (variantes):
                  </div>
                  <ul className="list-disc pl-5 space-y-0.5">
                    <li>
                      <b>sku</b> (sku, código, codigo, id, ean, upc)
                    </li>
                    <li>
                      <b>name</b> (name, producto, titulo, título, nombre)
                    </li>
                    <li>
                      <b>brand</b> (brand, marca)
                    </li>
                    <li>
                      <b>type</b> (type, categoria, categoría, rubro, familia,
                      departamento, línea)
                    </li>
                    <li>
                      <b>price</b> (price, precio, precio base, precio normal,
                      precio final)
                    </li>
                    <li>
                      <b>oldPrice</b> (oldPrice, precio normal, precio
                      referencia, precio anterior)
                    </li>
                    <li>
                      <b>description</b> (descripcion, descripción, detalle,
                      ficha)
                    </li>
                    <li>
                      <b>tags</b> (tags, etiquetas, keywords, palabras clave)
                    </li>
                    <li>
                      <b>image</b> (image, imagen, img, foto, url imagen, url)
                    </li>
                    <li>
                      <b>mostrar_en_tienda</b> (mostrar en tienda, visible,
                      publicado, mostrar, activo)
                    </li>
                  </ul>
                </div>
              </div>
            </div>
          </div>
        )}

        {detailOpen && detailProd && (
          <div className="fixed inset-0 z-[80]">
            <button
              className="absolute inset-0 bg-black/30"
              onClick={() => setDetailOpen(false)}
              aria-label="Cerrar"
            />
            <div className="absolute left-1/2 -translate-x-1/2 bottom-0 w-full max-w-[430px]">
              <div className="m-0 mx-auto rounded-t-3xl bg-white shadow-2xl border-t border-slate-200 overflow-hidden">
                <div className="flex items-center justify-between px-4 py-3 border-b border-slate-100">
                  <div className="font-medium">Detalle del producto</div>
                  <button
                    className="text-slate-500"
                    onClick={() => setDetailOpen(false)}
                  >
                    ✕
                  </button>
                </div>
                <div className="p-4 space-y-3">
                  <img
                    src={
                      isHttpUrl(detailProd.image)
                        ? detailProd.image
                        : FALLBACK_IMAGE
                    }
                    alt={detailProd.name}
                    className="w-full h-48 object-cover rounded-2xl bg-slate-50"
                  />
                  <div className="flex items-center justify-between">
                    <div>
                      <div className="text-sm text-slate-500">
                        {detailProd.brand || "(sin marca)"} ·{" "}
                        {detailProd.type || "(sin categoría)"}
                      </div>
                      <div className="text-lg font-semibold leading-snug">
                        {detailProd.name}
                      </div>
                    </div>
                    <div className="text-right">
                      <div className="text-lg font-semibold">
                        ${(detailProd.price || 0).toLocaleString("es-CL")}
                      </div>
                      {typeof detailProd.oldPrice === "number" &&
                        (detailProd.price || 0) > 0 &&
                        detailProd.oldPrice > (detailProd.price || 0) && (
                          <div className="text-xs text-slate-400 line-through">
                            ${detailProd.oldPrice.toLocaleString("es-CL")}
                          </div>
                        )}
                    </div>
                  </div>
                  {detailProd.badges && detailProd.badges.length > 0 && (
                    <div className="flex flex-wrap gap-2">
                      {detailProd.badges.map((b) => (
                        <span
                          key={b}
                          className={`text-[10px] px-2 py-0.5 rounded-full border ${badgeClass(
                            b
                          )}`}
                        >
                          {b}
                        </span>
                      ))}
                    </div>
                  )}
                  {detailProd.description && (
                    <p className="text-sm text-slate-700 whitespace-pre-wrap">
                      {detailProd.description}
                    </p>
                  )}
                  {detailProd.tags && detailProd.tags.length > 0 && (
                    <div className="text-[11px] text-slate-500 flex flex-wrap gap-1">
                      {detailProd.tags.slice(0, 12).map((t) => (
                        <span key={t}>#{t}</span>
                      ))}
                    </div>
                  )}
                  <div className="flex justify-end gap-2 pt-2">
                    <button
                      className="px-3 py-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-50"
                      onClick={() => toggleFavorite(detailProd.sku)}
                    >
                      {favorites.has(detailProd.sku)
                        ? "★ Quitar de fav"
                        : "☆ Fav"}
                    </button>
                    <button
                      className="px-3 py-2 rounded-xl bg-slate-900 text-white hover:bg-slate-800"
                      onClick={() => {
                        addToCart(detailProd);
                        setDetailOpen(false);
                      }}
                    >
                      Agregar
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* Botón flotante de voz */}
        <div className="pointer-events-none fixed bottom-24 right-4 left-0 flex justify-end z-[55]">
          <div className="pointer-events-auto flex flex-col items-end gap-2 mr-4">
            {lastTranscript && (
              <div className="max-w-[260px] text-[11px] bg-slate-900 text-white px-3 py-1.5 rounded-2xl shadow-lg">
                🎙️ {lastTranscript}
              </div>
            )}
            <button
              onClick={toggleListening}
              className={
                "w-11 h-11 rounded-full shadow-lg border flex items-center justify-center text-lg " +
                (listening
                  ? "bg-rose-600 text-white border-rose-700 animate-pulse"
                  : hasSpeechApi
                  ? "bg-slate-900 text-white border-slate-900"
                  : "bg-slate-200 text-slate-500 border-slate-300")
              }
              title={
                hasSpeechApi
                  ? listening
                    ? "Detener escucha"
                    : "Hablar para controlar la app"
                  : "La API de voz no está disponible en este navegador"
              }
            >
              {listening ? "⏹" : "🎙️"}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
// ==== FIN CÓDIGO APP NESTLÉ (TypeScript) ====
