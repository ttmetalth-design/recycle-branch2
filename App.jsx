import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Package, Users, FileText, ShoppingCart, Boxes, Plus, Trash2, Edit2,
  Search, Download, X, ChevronRight, ChevronLeft, ChevronDown, Menu, ArrowDownToLine,
  ArrowUpFromLine, History, TrendingUp, Save, Printer, Landmark,
  CheckCircle2, XCircle, Clock, CreditCard, PackageMinus, ArrowRight, Wallet, Receipt,
  Image, FileSpreadsheet, FileDown, Truck, Check,
  LayoutDashboard, ShoppingBag, BarChart3, BadgeDollarSign, ArrowLeftRight,
  Building2, ScrollText, PieChart, Settings, Tag, ClipboardList, Banknote
} from "lucide-react";
import { isSupabaseReady } from './supabase'
import { useSupabaseSync, loadAllFromSupabase, useSyncStatus, saveToSupabase } from './useSupabaseSync'
import { loadProducts, insertProduct, updateProduct, deleteProduct, useProductsRealtime } from './useProductsSync'
// ---------- Seed data ----------
const initialProducts = [];
 // ลูกค้าแต่ละคนสามารถมีบัญชีธนาคารได้หลายบัญชี: bankAccounts = [{id, bankName, accountNo, accountName}]
const initialCustomers = [];
  // บัญชีธนาคารของร้าน (store's own bank accounts)
const initialStoreBankAccounts = [];
 // payments = [{id, date, amount, fromStoreBankId ("CASH" หรือ id บัญชีร้าน), method}]
// receivingCustomerBankId = บัญชีลูกค้าที่จะรับเงิน (เลือกครั้งเดียวต่อใบ)
const initialPurchases = [];
const initialSales = [];
  // เบิกสินค้าเพื่อขาย: เบิกสินค้าต้นทางออกจากสต๊อก (ตัดสต๊อกทันทีตามต้นทุน FIFO)
// แล้วนำยอด (จำนวน + มูลค่า) ไปรวมเป็นต้นทุนของสินค้าเป้าหมายในใบขาย (sales invoice) ที่ระบุ
// value/avgCost คำนวณจากต้นทุน FIFO ของสต๊อกต้นทาง ณ ตอนที่บันทึกการเบิก
const initialWithdrawals = [];

// เงินมัดจำจ่ายให้ลูกค้าล่วงหน้า: {id, date, customerId, amount, note, fromStoreBankId}
// ยอดมัดจำคงเหลือของลูกค้า = ผลรวมเงินมัดจำที่จ่าย - ผลรวมเงินมัดจำที่ถูกหักในใบรับสินค้า (payments ที่ fromStoreBankId === "DEPOSIT")
const initialDeposits = [];

// ค่าใช้จ่าย: {id, date, category, description, amount, fromStoreBankId ("CASH" หรือ id บัญชีร้าน)}
const initialExpenses = [];

// เงินกู้ยืม / เช่าซื้อ: {id, name, type, principal, annualInterestRate, totalInstallments, startDate, lender}
// งวดผ่อนสร้างจาก amortization schedule (ผ่อนเท่ากันทุกเดือน ลดดอกเบี้ยจากเงินต้นคงเหลือ)
const initialLoans = [];
const initialAssets = [
  { id: "AS001", name: "รถกระบะบรรทุก", category: "ยานพาหนะ", purchaseDate: "2024-01-15", cost: 650000, lifeYears: 5, depreciationMethod: "เส้นตรง", note: "" },
  { id: "AS002", name: "เครื่องชั่งน้ำหนัก", category: "เครื่องจักร/อุปกรณ์", purchaseDate: "2024-03-01", cost: 45000, lifeYears: 10, depreciationMethod: "เส้นตรง", note: "" },
];
const initialShareholders = [
  { id: "SH1", name: "หุ้นส่วน 1", percent: 50 },
  { id: "SH2", name: "หุ้นส่วน 2", percent: 50 },
];
const LOAN_TYPES = ["เงินกู้ยืม", "เช่าซื้อ"];

// หมวดหมู่ใหญ่ (เพิ่มได้) และหมวดหมู่ย่อยเริ่มต้นของแต่ละหมวดหมู่ใหญ่ (เพิ่มได้)
const EXPENSE_MAIN_CATEGORIES = ["ค่าใช้จ่าย", "ภาษี", "สินทรัพย์", "สินเชื่อ"];
const EXPENSE_SUBCATEGORIES_DEFAULT = {
  "ค่าใช้จ่าย": ["ค่าน้ำมัน/ขนส่ง", "ค่าแรงงาน", "ค่าเช่า", "ค่าน้ำ/ค่าไฟ", "ค่าซ่อมบำรุง", "ค่าอุปกรณ์/วัสดุสิ้นเปลือง", "อื่นๆ"],
  "ภาษี": ["ภาษีมูลค่าเพิ่ม", "ภาษีเงินได้", "ภาษีหัก ณ ที่จ่าย", "ภาษีป้าย", "อื่นๆ"],
  "สินทรัพย์": ["ซื้อเครื่องจักร/อุปกรณ์", "ซื้อยานพาหนะ", "ซ่อมแซมปรับปรุง", "อื่นๆ"],
  "สินเชื่อ": ["ชำระเงินกู้ (เงินต้น)", "ชำระดอกเบี้ย", "ค่าธรรมเนียมสินเชื่อ", "อื่นๆ"],
};

const UNIT_OPTIONS_DEFAULT = ["กก.", "ตัน", "ชิ้น", "ม้วน", "ใบ"];
const PRODUCT_TYPES = ["กระดาษ", "พลาสติก", "เหล็ก", "อลูมิเนียม", "ทองแดง", "อื่นๆ"];
const PAYMENT_METHODS = ["เงินสด", "โอนเงิน", "เช็ค", "พร้อมเพย์"];
const MONTH_NAMES_TH = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
// ช่องทางชำระเงินสำหรับใบรับสินค้า (ไม่มีเงินสด เพราะต้องระบุบัญชีลูกค้าที่รับเงิน)
const PURCHASE_PAYMENT_CHANNELS = ["เงินสด", "โอนเงิน", "เช็ค", "พร้อมเพย์"];
const PAYMENT_STATUSES = ["รอชำระ", "ชำระแล้ว", "ชำระบางส่วน"];
const PURCHASE_STATUSES = ["รออนุมัติ", "อนุมัติแล้ว", "ยกเลิก"];
const BANK_NAMES = ["กสิกรไทย", "ไทยพาณิชย์", "กรุงไทย", "กรุงเทพ", "ทหารไทยธนชาต", "กรุงศรีอยุธยา", "ออมสิน", "ธ.ก.ส.", "ซีไอเอ็มบี", "ยูโอบี", "อื่นๆ"];

const fmt = (n) => Number(n || 0).toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtInt = (n) => Number(n || 0).toLocaleString("th-TH", { maximumFractionDigits: 2 });
// ปัดเศษสตางค์ของยอดเงินให้เป็นจำนวนเต็มบาท — ใช้ตอนกรอกยอดจ่าย/รับเงินจริง
const roundUpAmount = (n) => Math.ceil(Number(n) || 0);
const roundDownAmount = (n) => Math.floor(Number(n) || 0);

// ---------- Export Utilities ----------
// Download Excel (.xlsx) from a 2D array of rows
function exportExcel(rows, filename = "export.xlsx", sheetName = "Sheet1") {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

// Download PNG from a DOM element (id) using Canvas API
function exportImage(elementId, filename = "export.png") {
  const el = document.getElementById(elementId);
  if (!el) return;
  import("https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js")
    .then(() => { /* no-op, not available */ })
    .catch(() => {});
  // Fallback: use browser print with page setup
  const printWindow = window.open("", "_blank");
  if (!printWindow) { alert("กรุณาอนุญาต popup เพื่อบันทึกรูปภาพ"); return; }
  printWindow.document.write(`
    <html><head><title>${filename}</title>
    <style>body{margin:0;padding:20px;font-family:'Noto Sans Thai',sans-serif}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:6px 10px;font-size:12px}tr:nth-child(even){background:#f9f9f9}</style>
    </head><body>${el.innerHTML}</body></html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
  printWindow.close();
}

// Print current page as PDF (via browser print dialog)
// ---------- Print preview: global subscriber pattern ----------
// printAsPDF ถูกเรียกจากหลายจุดในแอปที่ไม่มี prop เข้าถึง state ของ App โดยตรง
// จึงใช้ตัวแปร global เก็บ callback ที่ App ลงทะเบียนไว้ตอน mount แทน
let __printPreviewSetter = null;
function registerPrintPreview(setter) { __printPreviewSetter = setter; }

function printAsPDF(elementId, title = "") {
  const el = document.getElementById(elementId);
  if (!el) { window.print(); return; }

  // มือถือ → ใช้ overlay (ไม่ต้องเปิด tab ใหม่ กลับแอพได้ทันที)
  // คอม → เปิดหน้าต่างใหม่ (พิมพ์/PDF สะดวกกว่า)
  const isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent);

  if (isMobile) {
    if (__printPreviewSetter) { __printPreviewSetter({ html: el.innerHTML, title }); }
    return;
  }

  const win = window.open('', '_blank', 'width=900,height=700');
  if (!win) {
    if (__printPreviewSetter) { __printPreviewSetter({ html: el.innerHTML, title }); }
    return;
  }

  const isLandscape = elementId === "transfer-sheet-print";
  win.document.write(`<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>${title}</title>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;600;700&display=swap">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: #e5e7eb; font-family: 'Noto Sans Thai', sans-serif; }
    .toolbar { position: sticky; top: 0; background: #1f2937; color: #fff; padding: 10px 20px; display: flex; align-items: center; justify-content: space-between; z-index: 100; gap: 12px; }
    .toolbar span { font-size: 14px; font-weight: 600; flex: 1; }
    .toolbar .zoom-controls { display: flex; align-items: center; gap: 6px; }
    .toolbar .zoom-controls button { background: #374151; color: #fff; border: none; width: 32px; height: 32px; border-radius: 6px; font-size: 18px; cursor: pointer; display: flex; align-items: center; justify-content: center; line-height: 1; }
    .toolbar .zoom-controls span { font-size: 13px; min-width: 44px; text-align: center; }
    .toolbar .print-btn { background: #0f6e56; color: #fff; border: none; padding: 8px 18px; border-radius: 6px; font-size: 13px; cursor: pointer; font-family: 'Noto Sans Thai', sans-serif; white-space: nowrap; }
    .page-wrap { padding: 20px 16px; display: flex; justify-content: center; }
    .page { background: #fff; width: 210mm; padding: 10mm; box-shadow: 0 2px 16px rgba(0,0,0,0.15); font-size: 11px; color: #1f2937; transform-origin: top center; transition: transform 0.15s; }
    table { border-collapse: collapse; width: 100%; page-break-inside: auto; }
    td, th { border: 1px solid #ddd; padding: 4px 6px; font-size: 10px; }
    th { background: #f3f4f6; font-weight: 700; }
    tr:nth-child(even) { background: #f9f9f9; }
    tfoot td { font-weight: 700; background: #f3f4f6; border-top: 2px solid #5a1414; }
    img { max-width: 100%; }
    button { display: none !important; }
    .toolbar button, .toolbar .print-btn { display: flex !important; }
    thead { display: table-header-group; }
    tfoot { display: table-footer-group; }
    tr { page-break-inside: avoid; page-break-after: auto; }
    tr, td, th { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
    ${isLandscape ? "@page { size: A4 landscape; margin: 6mm; }" : "@page { size: A4 portrait; margin: 10mm; }"}
    ${isLandscape ? ".page { width: 297mm !important; font-size: 9px !important; }" : ""}
    ${isLandscape ? "td, th { padding: 3px 5px !important; font-size: 9px !important; }" : ""}
    @media print {
      .toolbar { display: none !important; }
      body { background: #fff; }
      .page-wrap { padding: 0; }
      .page { width: 100%; margin: 0; padding: 0; box-shadow: none; transform: none !important; }
    }
  </style>
</head>
<body>
  <div class="toolbar">
    <span>${title}</span>
    <div class="zoom-controls">
      <button onclick="zoom(-10)">−</button>
      <span id="zoom-label">50%</span>
      <button onclick="zoom(+10)">+</button>
    </div>
    <button class="print-btn" onclick="window.print()">🖨️ พิมพ์ / บันทึก PDF</button>
  </div>
  <div class="page-wrap">
    <div class="page" id="page">${el.innerHTML}</div>
  </div>
  <script>
    var scale = 50;
    function zoom(delta) {
      scale = Math.min(150, Math.max(50, scale + delta));
      document.getElementById('page').style.transform = 'scale(' + scale/100 + ')';
      document.getElementById('page').style.marginBottom = scale < 100 ? ((scale - 100) * 2.97) + 'mm' : '0';
      document.getElementById('zoom-label').textContent = scale + '%';
    }
    zoom(0);
  <\/script>
</body>
</html>`);
  win.document.close();
}

// หน้าต่างดูตัวอย่างเอกสารแบบเต็มจอ ก่อนสั่งพิมพ์จริง
function PrintPreviewOverlay({ preview, onClose }) {
  if (!preview) return null;
  const [zoom, setZoom] = React.useState(50);

  const handlePrint = () => {
    const isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent);
    if (isMobile) { window.print(); return; }
    const win = window.open('', '_blank', 'width=900,height=700');
    if (!win) { window.print(); return; }
    win.document.write(`<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${preview.title || 'เอกสาร'}</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;600;700&display=swap">
<style>
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Noto Sans Thai', sans-serif; font-size: 11px; color: #1f2937; background: #fff; padding: 10mm; }
table { border-collapse: collapse; width: 100%; page-break-inside: auto; table-layout: fixed; }
td, th { border: 1px solid #ddd; padding: 4px 6px; font-size: 10px; word-break: break-word; overflow-wrap: break-word; }
th { background: #f3f4f6; font-weight: 700; }
tr:nth-child(even) { background: #f9f9f9; }
tfoot td { font-weight: 700; background: #f3f4f6; border-top: 2px solid #5a1414; }
img { max-width: 100%; }
button { display: none !important; }
thead { display: table-header-group; }
tfoot { display: table-footer-group; }
tr { page-break-inside: avoid; page-break-after: auto; }
tr, td, th { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
@page { size: A4 landscape; margin: 8mm; }
@media print { body { padding: 0; } }
</style></head><body>${preview.html}</body></html>`);
    win.document.close();
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "#fff", zIndex: 9999, display: "flex", flexDirection: "column" }}>
      {/* Toolbar */}
      <div style={{ display: "flex", alignItems: "center", padding: "8px 12px", background: "#1f2937", flexShrink: 0, gap: 8 }}>
        <button style={{ ...btnSecondary, fontSize: 12, padding: "6px 10px", flexShrink: 0 }} onClick={onClose}><ChevronLeft size={14} /> ย้อนกลับ</button>
        <div style={{ flex: 1, fontWeight: 700, fontSize: 12, color: "#fff", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{preview.title}</div>
        <div style={{ display: "flex", alignItems: "center", gap: 4, flexShrink: 0 }}>
          <button onClick={() => setZoom(z => Math.max(30, z - 10))} style={{ background: "#374151", color: "#fff", border: "none", width: 30, height: 30, borderRadius: 6, fontSize: 18, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>−</button>
          <span style={{ color: "#fff", fontSize: 12, minWidth: 38, textAlign: "center" }}>{zoom}%</span>
          <button onClick={() => setZoom(z => Math.min(150, z + 10))} style={{ background: "#374151", color: "#fff", border: "none", width: 30, height: 30, borderRadius: 6, fontSize: 18, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>+</button>
        </div>
        <button style={{ ...btnPrimary, fontSize: 12, padding: "6px 10px", flexShrink: 0 }} onClick={handlePrint}><Download size={14} /> PDF</button>
      </div>
      {/* Content */}
      <div style={{ flex: 1, overflow: "auto", background: "#e5e7eb", WebkitOverflowScrolling: "touch" }}>
        <div style={{ display: "flex", justifyContent: "center", padding: "12px 8px", minWidth: "max-content" }}>
          <div
            id="print-preview-content"
            style={{ background: "#fff", width: "210mm", padding: "8mm", boxShadow: "0 2px 12px rgba(0,0,0,0.12)", fontFamily: "'Noto Sans Thai', sans-serif", fontSize: 11, color: "#1f2937", transform: `scale(${zoom/100})`, transformOrigin: "top left", transition: "transform 0.15s" }}
            dangerouslySetInnerHTML={{ __html: preview.html }}
          />
        </div>
      </div>
      <style>{`
        #print-preview-content table { border-collapse: collapse; width: 100%; }
        #print-preview-content td, #print-preview-content th { border: 1px solid #ddd; padding: 4px 6px; font-size: 10px; }
        #print-preview-content th { background: #f3f4f6; font-weight: 700; }
        #print-preview-content tr:nth-child(even) { background: #f9f9f9; }
        #print-preview-content tfoot td { font-weight: 700; background: #f3f4f6; border-top: 2px solid #5a1414; }
        #print-preview-content img { max-width: 100%; }
        #print-preview-content button { display: none !important; }
        #print-preview-content tr, #print-preview-content td, #print-preview-content th {
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
          color-adjust: exact !important;
        }
      `}</style>
    </div>
  );
}

// ---------- Confirm dialog: global subscriber pattern (เหมือน printAsPDF) ----------
// confirmAction ถูกเรียกจากหลายจุดในแอปที่ไม่มี prop เข้าถึง state ของ App โดยตรง
let __confirmSetter = null;
function registerConfirmDialog(setter) { __confirmSetter = setter; }

function confirmAction(message, onConfirm, options = {}) {
  if (__confirmSetter) {
    __confirmSetter({ message, onConfirm, title: options.title, confirmLabel: options.confirmLabel });
  } else {
    // เผื่อกรณี dialog ยังไม่ได้ลงทะเบียน (ไม่ควรเกิดขึ้นในการใช้งานปกติ)
    if (window.confirm(message)) onConfirm();
  }
}

function ConfirmDialog({ pending, onClose }) {
  if (!pending) return null;
  return (
    <Modal title={pending.title || "ยืนยันการลบ"} onClose={onClose} zIndex={200}>
      <div style={{ display: "flex", alignItems: "flex-start", gap: 12, padding: "4px 0 16px" }}>
        <div style={{ width: 36, height: 36, borderRadius: "50%", background: "#faece7", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
          <Trash2 size={18} color="#993c1d" />
        </div>
        <p style={{ margin: 0, fontSize: 14, color: "#374151", lineHeight: 1.6 }}>{pending.message}</p>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
        <button style={btnSecondary} onClick={onClose}>ยกเลิก</button>
        <button style={btnDanger} onClick={() => { pending.onConfirm(); onClose(); }}>
          <Trash2 size={14} /> {pending.confirmLabel || "ยืนยันการลบ"}
        </button>
      </div>
    </Modal>
  );
}

// ---------- Keyboard navigation helper: Enter = move to next field, last field = submit ----------
function handleEnterNavigate(e, onSubmit) {
  if (e.key !== "Enter") return;
  // อย่าไปยุ่งกับ textarea (ต้องกด Enter ขึ้นบรรทัดใหม่ได้ตามปกติ)
  if (e.target.tagName === "TEXTAREA") return;
  e.preventDefault();

  const form = e.target.closest('[data-kbform]');
  if (!form) return;

  const focusable = Array.from(
    form.querySelectorAll('input:not([disabled]):not([type="hidden"]), select:not([disabled]), textarea:not([disabled]), button[data-kbsubmit]')
  ).filter((el) => el.offsetParent !== null); // เฉพาะที่มองเห็นอยู่ (ไม่ถูกซ่อน)

  const idx = focusable.indexOf(e.target);
  if (idx === -1) return;

  const next = focusable[idx + 1];
  if (next) {
    if (next.tagName === "BUTTON" && next.hasAttribute("data-kbsubmit")) {
      // ถึงปุ่มบันทึกแล้ว ให้กดบันทึกเลย
      next.click();
    } else {
      next.focus();
      if (next.select) next.select();
    }
  } else if (onSubmit) {
    onSubmit();
  }
}
// ExportToolbar component — renders 3 export buttons for a section
function ExportToolbar({ onPDF, onExcel, onImage, label = "" }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
      {label && <span style={{ fontSize: 12, color: "#6b7280", marginRight: 4 }}>{label}</span>}
      <button
        onClick={onPDF}
        title="บันทึก PDF"
        style={{ display: "flex", alignItems: "center", gap: 4, padding: "5px 10px", borderRadius: 6, border: "1px solid #e5e7eb", background: "#fff", cursor: "pointer", fontSize: 12, color: "#993c1d" }}
      >
        <FileDown size={13} /> PDF
      </button>
      <button
        onClick={onExcel}
        title="บันทึก Excel"
        style={{ display: "flex", alignItems: "center", gap: 4, padding: "5px 10px", borderRadius: 6, border: "1px solid #e5e7eb", background: "#fff", cursor: "pointer", fontSize: 12, color: "#0f6e56" }}
      >
        <FileSpreadsheet size={13} /> Excel
      </button>
      <button
        onClick={onImage}
        title="บันทึกรูปภาพ"
        style={{ display: "flex", alignItems: "center", gap: 4, padding: "5px 10px", borderRadius: 6, border: "1px solid #e5e7eb", background: "#fff", cursor: "pointer", fontSize: 12, color: "#185fa5" }}
      >
        <Image size={13} /> รูปภาพ
      </button>
    </div>
  );
}


// Builds inventory lots, consumes FIFO on sales/withdrawals, returns stock summary + movement history
function computeInventory(products, purchases, sales, withdrawals = []) {
  const lots = {};
  const movements = [];

  products.forEach((p) => (lots[p.id] = []));

  const events = [];

  // ===== ยอดยกมา: ใส่เป็น event ณ วันที่ 1 ของเดือนที่ระบุไว้ (ถ้าไม่ระบุเดือน ใช้วันที่เริ่มต้นสุดสำหรับความเข้ากันได้กับข้อมูลเดิม) =====
  products.forEach((p) => {
    const qty  = Number(p.openingQty)  || 0;
    const cost = Number(p.openingCost) || 0;
    if (qty > 0) {
      const openingDate = p.openingMonth ? `${p.openingMonth}-01` : "0000-01-01";
      events.push({ type: "in", date: openingDate, ref: "ยอดยกมา", productId: p.id, qty, price: cost, isOpening: true });
    }
  });

  purchases.forEach((po) => {
    if (po.status !== "อนุมัติแล้ว") return;
    po.items.forEach((it) => {
      events.push({ type: "in", date: po.date, ref: po.id, productId: it.productId, qty: it.net, price: it.price });
    });
  });
  sales.forEach((inv) => {
    inv.items.forEach((it) => {
      if (it.fromWithdrawal) return; // สต๊อกถูกตัดไปแล้วตอนเบิก ไม่ต้องตัดซ้ำที่นี่
      events.push({ type: "out", date: inv.date, ref: inv.id, productId: it.productId, qty: it.net });
    });
  });
  withdrawals.forEach((lot) => {
    (lot.items || []).forEach((it) => {
      events.push({ type: "withdraw", date: lot.date, ref: lot.id, productId: it.sourceProductId, qty: it.qty });
    });
  });
  // เรียงตามวันที่ แล้วให้ "withdraw" มาก่อน "in"/"out" ในวันเดียวกัน เพื่อให้ลำดับสอดคล้องกับการตัดสต๊อกทันที
  const typeOrder = { in: 0, withdraw: 1, out: 2 };
  events.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : (typeOrder[a.type] ?? 1) - (typeOrder[b.type] ?? 1)));

  events.forEach((ev) => {
    if (!lots[ev.productId]) lots[ev.productId] = [];
    if (ev.type === "in") {
      lots[ev.productId].push({ date: ev.date, ref: ev.ref, qtyRemaining: ev.qty, qtyOriginal: ev.qty, unitCost: ev.price });
      movements.push({ ...ev, balanceQty: null });
    } else {
      let remainingToConsume = ev.qty;
      let costConsumed = 0;
      const queue = lots[ev.productId];
      for (let i = 0; i < queue.length && remainingToConsume > 0; i++) {
        const lot = queue[i];
        if (lot.qtyRemaining <= 0) continue;
        const take = Math.min(lot.qtyRemaining, remainingToConsume);
        lot.qtyRemaining -= take;
        costConsumed += take * lot.unitCost;
        remainingToConsume -= take;
      }
      const avgCostUsed = ev.qty > 0 ? costConsumed / ev.qty : 0;
      movements.push({ ...ev, costConsumed, avgCostUsed, shortfall: remainingToConsume });
    }
  });

  const summary = products.map((p) => {
    const remaining = (lots[p.id] || []).reduce((s, l) => s + Math.max(0, l.qtyRemaining), 0);
    const totalCost = (lots[p.id] || []).reduce((s, l) => s + Math.max(0, l.qtyRemaining) * l.unitCost, 0);
    const avgCost = remaining > 0 ? totalCost / remaining : 0;
    return { productId: p.id, name: p.name, unit: p.unit, qty: remaining, totalCost, avgCost };
  });

  // Build per-product movement history with running balance
  const history = {};
  products.forEach((p) => {
    let balance = 0;
    history[p.id] = events
      .filter((e) => e.productId === p.id)
      .map((e) => {
        if (e.type === "in") balance += e.qty;
        else balance -= e.qty;
        return { ...e, balance };
      });
  });

  return { summary, history, lots, movements };
}

// คำนวณต้นทุน FIFO ของจำนวนที่จะเบิก โดยอิงจากสต๊อกคงเหลือปัจจุบัน (ไม่แก้ไข lots จริง)
function computeWithdrawalCost(inventory, sourceProductId, qty) {
  const lots = (inventory.lots[sourceProductId] || []).map((l) => ({ ...l }));
  let remaining = Number(qty) || 0;
  let cost = 0;
  let shortfall = 0;
  for (let i = 0; i < lots.length && remaining > 0; i++) {
    const lot = lots[i];
    if (lot.qtyRemaining <= 0) continue;
    const take = Math.min(lot.qtyRemaining, remaining);
    cost += take * lot.unitCost;
    lot.qtyRemaining -= take;
    remaining -= take;
  }
  if (remaining > 0) {
    // ไม่พอในสต๊อก ให้ใช้ต้นทุนเฉลี่ยปัจจุบันสำหรับส่วนที่เกิน
    const summary = inventory.summary.find((s) => s.productId === sourceProductId);
    const fallbackCost = summary?.avgCost || 0;
    cost += remaining * fallbackCost;
    shortfall = remaining;
  }
  return { value: cost, shortfall };
}

// ---------- Deposit balance helper ----------
// คำนวณยอดมัดจำคงเหลือของลูกค้าแต่ละราย
// = ผลรวมเงินมัดจำที่จ่ายให้ลูกค้า (deposits) - ผลรวมเงินมัดจำที่ถูกหักในใบรับสินค้า
//   (purchases[].payments[] ที่ fromStoreBankId === "DEPOSIT")
function computeDepositBalances(customers, deposits, purchases) {
  const given = {}; // customerId -> total given
  const used = {}; // customerId -> total used in purchases
  deposits.forEach((d) => {
    given[d.customerId] = (given[d.customerId] || 0) + (Number(d.amount) || 0);
  });
  purchases.forEach((po) => {
    (po.payments || []).forEach((p) => {
      if (p.fromStoreBankId === "DEPOSIT") {
        used[po.customerId] = (used[po.customerId] || 0) + (Number(p.amount) || 0);
      }
    });
  });
  return customers.map((c) => {
    const opening = Number(c.depositOpening) || 0;
    const newGiven = given[c.id] || 0;
    const totalGiven = opening + newGiven;
    const totalUsed = used[c.id] || 0;
    return { customerId: c.id, name: c.name, opening, newGiven, totalGiven, totalUsed, remaining: totalGiven - totalUsed };
  });
}

// รับล่วงหน้า (ฝั่งขาย) — ลูกค้าจ่ายเงินให้ร้านก่อนรับสินค้า
function computePrepaymentBalances(customers, prepayments, sales) {
  const received = {}; // customerId -> รับมาแล้วทั้งหมด
  const used = {};     // customerId -> ถูกหักในใบขายแล้ว
  prepayments.forEach((p) => {
    received[p.customerId] = (received[p.customerId] || 0) + (Number(p.amount) || 0);
  });
  sales.forEach((inv) => {
    (inv.payments || []).forEach((p) => {
      if (p.fromStoreBankId === "PREPAYMENT" || p.toStoreBankId === "PREPAYMENT") {
        used[inv.customerId] = (used[inv.customerId] || 0) + (Number(p.amount) || 0);
      }
    });
  });
  return customers.map((c) => {
    const opening = Number(c.prepaymentOpening) || 0;
    const newReceived = received[c.id] || 0;
    const totalReceived = opening + newReceived;
    const totalUsed = used[c.id] || 0;
    return { customerId: c.id, name: c.name, opening, newReceived, totalReceived, totalUsed, remaining: totalReceived - totalUsed };
  });
}

// ---------- Loan amortization schedule ----------
// คำนวณตารางผ่อนชำระแบบผ่อนเท่ากันทุกเดือน (ลดดอกเบี้ยจากเงินต้นคงเหลือไปเรื่อยๆ)
// คืนค่า: array ของงวด {no, dueDate, payment, interest, principalPortion, remainingBalance}
// คำนวณวันครบกำหนดของงวดที่ i: เดือนของ (startDate + i) แต่ใช้วันที่ = dueDayOfMonth
// (ถ้า dueDayOfMonth เกินจำนวนวันในเดือนนั้น ให้ใช้วันสุดท้ายของเดือนแทน เช่น 31 ก.พ. -> 28/29 ก.พ.)
function computeDueDate(startDate, monthOffset, dueDayOfMonth) {
  const d = new Date(startDate);
  const targetMonth = d.getMonth() + monthOffset;
  const targetYear = d.getFullYear() + Math.floor(targetMonth / 12);
  const normalizedMonth = ((targetMonth % 12) + 12) % 12;
  const lastDayOfMonth = new Date(targetYear, normalizedMonth + 1, 0).getDate();
  const day = Math.min(Math.max(1, Number(dueDayOfMonth) || d.getDate()), lastDayOfMonth);
  const result = new Date(targetYear, normalizedMonth, day);
  return result.toISOString().slice(0, 10);
}

function computeAmortizationSchedule(loan) {
  const principal = Number(loan.principal) || 0;
  const n = Number(loan.totalInstallments) || 0;
  if (principal <= 0 || n <= 0) return [];

  const startDate = loan.startDate ? new Date(loan.startDate) : new Date();
  const schedule = [];

  if (loan.interestMode === "amount") {
    // ดอกเบี้ยกรอกเป็นจำนวนเงินรวมตลอดสัญญา -> กระจายดอกเบี้ยเท่าๆกันทุกงวด (เหมือนเช่าซื้อทั่วไป)
    const totalInterest = Number(loan.totalInterestAmount) || 0;
    const interestPerInstallment = totalInterest / n;
    const principalPerInstallment = principal / n;
    const payment = principalPerInstallment + interestPerInstallment;
    let balance = principal;
    for (let i = 1; i <= n; i++) {
      let principalPortion = principalPerInstallment;
      let thisInterest = interestPerInstallment;
      let thisPayment = payment;
      if (i === n) {
        // งวดสุดท้าย: ปรับให้พอดีกับเงินต้นคงเหลือ (กันเศษทศนิยมสะสม)
        principalPortion = balance;
        thisPayment = principalPortion + thisInterest;
      }
      balance = Math.max(0, balance - principalPortion);
      const dueDate = computeDueDate(startDate, i, loan.dueDayOfMonth);
      schedule.push({ no: i, dueDate, payment: thisPayment, interest: thisInterest, principalPortion, remainingBalance: balance });
    }
    return schedule;
  }

  // ดอกเบี้ยกรอกเป็น % ต่อปี -> สูตรผ่อนเท่ากันทุกเดือน (annuity, ลดต้นลดดอก)
  const annualRate = Number(loan.annualInterestRate) || 0;
  const monthlyRate = annualRate / 100 / 12;
  let payment;
  if (monthlyRate === 0) {
    payment = principal / n;
  } else {
    payment = (principal * monthlyRate) / (1 - Math.pow(1 + monthlyRate, -n));
  }

  let balance = principal;
  for (let i = 1; i <= n; i++) {
    const interest = balance * monthlyRate;
    let principalPortion = payment - interest;
    let thisPayment = payment;
    // งวดสุดท้าย: ปรับให้พอดีกับเงินต้นคงเหลือ (กันเศษทศนิยมสะสม)
    if (i === n) {
      principalPortion = balance;
      thisPayment = principalPortion + interest;
    }
    balance = Math.max(0, balance - principalPortion);

    const dueDate = computeDueDate(startDate, i, loan.dueDayOfMonth);

    schedule.push({
      no: i,
      dueDate,
      payment: thisPayment,
      interest,
      principalPortion,
      remainingBalance: balance,
    });
  }
  return schedule;
}

// ---------- Generic small UI bits ----------
function Modal({ title, onClose, children, wide, fullscreen, zIndex: zIndexProp }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: zIndexProp || 50, padding: fullscreen ? 0 : "1rem" }}>
      <div style={{ background: "var(--color-background-primary, #fff)", borderRadius: fullscreen ? 0 : 12, width: fullscreen ? "100vw" : wide ? "min(900px, 95vw)" : "min(520px, 95vw)", height: fullscreen ? "100vh" : undefined, maxHeight: fullscreen ? "100vh" : "90vh", overflowY: "auto", boxShadow: "0 8px 30px rgba(0,0,0,0.15)" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "1rem 1.25rem", borderBottom: "1px solid #e5e7eb" }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 600 }}>{title}</h3>
          <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer", padding: 4, color: "#6b7280" }}><X size={20} /></button>
        </div>
        <div style={{ padding: "1.25rem" }}>{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <label style={{ display: "block", fontSize: 13, fontWeight: 500, color: "#374151", marginBottom: 4 }}>{label}</label>
      {children}
    </div>
  );
}

function Header({ title, subtitle, children }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20, gap: 12, flexWrap: "wrap" }}>
      <div>
        <h2 style={{ margin: 0, fontSize: 20, fontWeight: 700 }}>{title}</h2>
        {subtitle && <p style={{ margin: "4px 0 0", fontSize: 13, color: "#6b7280" }}>{subtitle}</p>}
      </div>
      {children && <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>{children}</div>}
    </div>
  );
}

// NumInput — input ตัวเลขที่แสดงลูกน้ำขณะพิมพ์
// ตอน focus: แสดงตัวเลขล้วน (พิมพ์ได้ปกติ)
// ตอน blur: แสดงพร้อมลูกน้ำ เช่น 10,000.50
function NumInput({ value, onChange, onKeyDown, style, placeholder, min }) {
  const [focused, setFocused] = React.useState(false);
  const num = parseFloat(String(value).replace(/,/g, "")) || 0;
  const formatted = focused
    ? (value === 0 || value === "0" ? "" : String(value))
    : (num === 0 ? "0" : num.toLocaleString("en-US", { maximumFractionDigits: 4 }));

  return (
    <input
      type="text"
      inputMode="decimal"
      style={style}
      placeholder={placeholder}
      value={formatted}
      onFocus={(e) => { setFocused(true); e.target.select(); }}
      onBlur={() => setFocused(false)}
      onChange={(e) => {
        // อนุญาตเฉพาะตัวเลข จุดทศนิยม และลบหน้า
        const raw = e.target.value.replace(/[^0-9.]/g, "");
        onChange({ target: { value: raw } });
      }}
      onKeyDown={onKeyDown}
    />
  );
}

function SearchBar({ value, onChange, placeholder, dateFrom, dateTo, onDateFromChange, onDateToChange }) {
  return (
    <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
      <div style={{ position: "relative", flex: "1 1 200px", minWidth: 180 }}>
        <Search size={16} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: "#9ca3af" }} />
        <input style={{ ...inputStyle, paddingLeft: 32 }} value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder} />
      </div>
      {onDateFromChange && (
        <div style={{ display: "flex", alignItems: "center", gap: 6, flex: "0 0 auto" }}>
          <span style={{ fontSize: 12, color: "#6b7280", whiteSpace: "nowrap" }}>วันที่</span>
          <input type="date" style={{ ...inputStyle, width: 140 }} value={dateFrom || ""} onChange={(e) => onDateFromChange(e.target.value)} />
          <span style={{ fontSize: 12, color: "#6b7280" }}>ถึง</span>
          <input type="date" style={{ ...inputStyle, width: 140 }} value={dateTo || ""} onChange={(e) => onDateToChange(e.target.value)} />
          {(dateFrom || dateTo) && (
            <button style={{ ...btnSecondary, padding: "4px 8px", fontSize: 12 }} onClick={() => { onDateFromChange(""); onDateToChange(""); }}>
              <X size={12} /> ล้าง
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function Card({ children, style }) {
  return (
    <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "auto", ...style }}>
      {children}
    </div>
  );
}

const inputStyle = {
  width: "100%", padding: "8px 10px", borderRadius: 8, border: "1px solid #d1d5db",
  fontSize: 14, boxSizing: "border-box", fontFamily: "inherit",
};

const btnPrimary = {
  background: "#0f6e56", color: "#fff", border: "none", padding: "8px 16px",
  borderRadius: 8, fontSize: 14, fontWeight: 500, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6,
};
const btnSecondary = {
  background: "#fff", color: "#374151", border: "1px solid #d1d5db", padding: "8px 16px",
  borderRadius: 8, fontSize: 14, fontWeight: 500, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6,
};
const btnDanger = {
  background: "#fff", color: "#a32d2d", border: "1px solid #f7c1c1", padding: "6px 10px",
  borderRadius: 8, fontSize: 13, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4,
};
const iconBtn = {
  background: "#fff", border: "1px solid #d1d5db", padding: "6px 10px", borderRadius: 8,
  cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 4, fontSize: 13, color: "#374151",
};
const roundBtn = {
  background: "#eef6f2", border: "1.5px solid #1d9e75", padding: "9px 11px", borderRadius: 8,
  cursor: "pointer", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 700, color: "#0f6e56", lineHeight: 1, flexShrink: 0, minWidth: 38,
};

const thStyle = { textAlign: "left", padding: "10px 12px", fontSize: 12, fontWeight: 600, color: "#6b7280", borderBottom: "1px solid #e5e7eb", whiteSpace: "nowrap" };
const tdStyle = { padding: "10px 12px", fontSize: 14, borderBottom: "1px solid #f3f4f6", whiteSpace: "nowrap" };

function genId(prefix, list, dateStr) {
  const now = dateStr ? new Date(dateStr + "T00:00:00") : new Date();
  const yy = now.getFullYear().toString().slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const datePart = `${yy}${mm}${dd}`;
  const dayPrefix = `${prefix}${datePart}`;
  const dayCount = list.filter((item) => {
    const id = typeof item === "string" ? item : (item.refNo || item.id || "");
    return id.startsWith(dayPrefix);
  }).length;
  return `${dayPrefix}${String(dayCount + 1).padStart(3, "0")}`;
}
function genSeqId(prefix, list) {
  const nums = list
    .map((item) => {
      const id = typeof item === "string" ? item : (item.id || "");
      const match = id.match(new RegExp(`^${prefix}(\\d+)$`));
      return match ? parseInt(match[1]) : 0;
    })
    .filter((n) => n > 0);
  const next = nums.length > 0 ? Math.max(...nums) + 1 : 1;
  return `${prefix}${String(next).padStart(3, "0")}`;
}

// ---------- Searchable product select (type to filter) ----------
function ProductSelect({ products, value, onChange, disabled, minWidth = 170, labelWithId = true }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [rect, setRect] = useState(null);
  const [highlight, setHighlight] = useState(0); // index ของรายการที่ไฮไลท์ไว้ด้วยคีย์บอร์ด
  const wrapRef = React.useRef(null);
  const inputRef = React.useRef(null);
  const listRef = React.useRef(null);

  const selected = products.find((p) => p.id === value);
  const display = (p) => (labelWithId ? `${p.id} · ${p.name}` : p.name);

  const filtered = query.trim()
    ? products.filter((p) => p.id.toLowerCase().includes(query.toLowerCase()) || p.name.toLowerCase().includes(query.toLowerCase()))
    : products;

  const updateRect = () => {
    if (inputRef.current) {
      const r = inputRef.current.getBoundingClientRect();
      setRect({ top: r.bottom, left: r.left, width: r.width });
    }
  };

  React.useEffect(() => {
    if (!open) return;
    updateRect();
    setHighlight(0);
    const handler = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target) && !(e.target.closest && e.target.closest('[data-product-select-dropdown]'))) {
        setOpen(false);
        setQuery("");
      }
    };
    const reposition = () => updateRect();
    document.addEventListener("mousedown", handler);
    window.addEventListener("scroll", reposition, true);
    window.addEventListener("resize", reposition);
    return () => {
      document.removeEventListener("mousedown", handler);
      window.removeEventListener("scroll", reposition, true);
      window.removeEventListener("resize", reposition);
    };
  }, [open]);

  // เลื่อนรายการที่ไฮไลท์ให้อยู่ในมุมมองเสมอ
  React.useEffect(() => {
    if (!open || !listRef.current) return;
    const el = listRef.current.querySelector(`[data-idx="${highlight}"]`);
    if (el) el.scrollIntoView({ block: "nearest" });
  }, [highlight, open]);

  const selectAt = (idx) => {
    const p = filtered[idx];
    if (!p) return;
    onChange(p.id);
    setOpen(false);
    setQuery("");
  };

  // ปุ่มลูกศรขึ้น/ลง = เลื่อนไฮไลท์, Enter = เลือก, Escape = ปิด
  const handleKeyDown = (e) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      if (!open) { setOpen(true); return; }
      setHighlight((h) => Math.min(h + 1, filtered.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      if (!open) { setOpen(true); return; }
      setHighlight((h) => Math.max(h - 1, 0));
    } else if (e.key === "Enter") {
      if (open && filtered.length > 0) {
        e.preventDefault();
        e.stopPropagation();
        selectAt(highlight);
        // หลังเลือกแล้ว ส่ง Enter ต่อไปยัง logic เลื่อนช่องถัดไปของฟอร์ม (ถ้ามี)
        setTimeout(() => {
          const fakeEvent = { key: "Enter", target: inputRef.current, preventDefault: () => {} };
          if (typeof handleEnterNavigate === "function") handleEnterNavigate(fakeEvent);
        }, 0);
      }
    } else if (e.key === "Escape") {
      setOpen(false);
      setQuery("");
    }
  };

  const dropdown = open && rect && (
    <div
      ref={listRef}
      data-product-select-dropdown
      style={{
        position: "fixed", top: rect.top, left: rect.left, width: rect.width, zIndex: 1000,
        background: "#fff", border: "1px solid #d1d5db", borderRadius: 8,
        marginTop: 4, maxHeight: 220, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,0.08)",
      }}
    >
      {filtered.length === 0 ? (
  <div
    style={{
      padding: "8px 10px",
      fontSize: 13,
      color: "#9ca3af",
    }}
  >
    ไม่พบสินค้า
  </div>
) : (
  filtered.map((p, idx) => (
    <div
      key={p.id}
      data-idx={idx}
      onMouseDown={(e) => {
        e.preventDefault();
        selectAt(idx);
      }}
      onMouseEnter={() => setHighlight(idx)}
      style={{
        padding: "8px 10px",
        fontSize: 13,
        cursor: "pointer",
        background: idx === highlight ? "#e0ddfb" : (p.id === value ? "#eeedfe" : "#fff"),
      }}
    >
      {display(p)}
        </div>
      ))
    )}
    </div>
  );  

  return (
    <div ref={wrapRef} style={{ position: "relative", minWidth }}>
      <input
        ref={inputRef}
        style={{ ...inputStyle, paddingRight: selected && !open ? 28 : undefined }}
        disabled={disabled}
        placeholder="ค้นหาสินค้า..."
        value={open ? query : (selected ? display(selected) : "")}
        onFocus={() => { setOpen(true); setQuery(""); }}
        onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
        onKeyDown={handleKeyDown}
      />
      {selected && !open && !disabled && (
        <button
          type="button"
          onClick={() => onChange("")}
          title="ล้างค่าที่เลือก"
          style={{
            position: "absolute", right: 6, top: "50%", transform: "translateY(-50%)",
            background: "none", border: "none", cursor: "pointer", color: "#9ca3af",
            display: "flex", alignItems: "center", padding: 2,
          }}
        >
          <X size={14} />
        </button>
      )}
      {dropdown}
    </div>
  );
}

// ---------- Searchable customer select (type to filter) ----------
function CustomerSelect({ customers, value, onChange, disabled, minWidth = 180, labelWithId = true }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [rect, setRect] = useState(null);
  const [highlight, setHighlight] = useState(0);
  const wrapRef = React.useRef(null);
  const inputRef = React.useRef(null);
  const listRef = React.useRef(null);

  const selected = customers.find((c) => c.id === value);
  const display = (c) => (labelWithId ? `${c.id} · ${c.name}` : c.name);

  const filtered = query.trim()
    ? customers.filter((c) => c.id.toLowerCase().includes(query.toLowerCase()) || c.name.toLowerCase().includes(query.toLowerCase()))
    : customers;

  const updateRect = () => {
    if (inputRef.current) {
      const r = inputRef.current.getBoundingClientRect();
      setRect({ top: r.bottom, left: r.left, width: r.width });
    }
  };

  React.useEffect(() => {
    if (!open) return;
    updateRect();
    setHighlight(0);
    const handler = (e) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target) && !(e.target.closest && e.target.closest('[data-customer-select-dropdown]'))) {
        setOpen(false);
        setQuery("");
      }
    };
    const reposition = () => updateRect();
    document.addEventListener("mousedown", handler);
    window.addEventListener("scroll", reposition, true);
    window.addEventListener("resize", reposition);
    return () => {
      document.removeEventListener("mousedown", handler);
      window.removeEventListener("scroll", reposition, true);
      window.removeEventListener("resize", reposition);
    };
  }, [open]);

  React.useEffect(() => {
    if (!open || !listRef.current) return;
    const el = listRef.current.querySelector(`[data-idx="${highlight}"]`);
    if (el) el.scrollIntoView({ block: "nearest" });
  }, [highlight, open]);

  const selectAt = (idx) => {
    const c = filtered[idx];
    if (!c) return;
    onChange(c.id);
    setOpen(false);
    setQuery("");
  };

  const handleKeyDown = (e) => {
    if (e.key === "ArrowDown") {
      e.preventDefault();
      if (!open) { setOpen(true); return; }
      setHighlight((h) => Math.min(h + 1, filtered.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      if (!open) { setOpen(true); return; }
      setHighlight((h) => Math.max(h - 1, 0));
    } else if (e.key === "Enter") {
      if (open && filtered.length > 0) {
        e.preventDefault();
        e.stopPropagation();
        selectAt(highlight);
        setTimeout(() => {
          const fakeEvent = { key: "Enter", target: inputRef.current, preventDefault: () => {} };
          if (typeof handleEnterNavigate === "function") handleEnterNavigate(fakeEvent);
        }, 0);
      }
    } else if (e.key === "Escape") {
      setOpen(false);
      setQuery("");
    }
  };

  const dropdown = open && rect && (
    <div
      ref={listRef}
      data-customer-select-dropdown
      style={{
        position: "fixed", top: rect.top, left: rect.left, width: rect.width, zIndex: 1000,
        background: "#fff", border: "1px solid #d1d5db", borderRadius: 8,
        marginTop: 4, maxHeight: 220, overflowY: "auto", boxShadow: "0 4px 12px rgba(0,0,0,0.08)",
      }}
    >
      {filtered.length === 0 && <div style={{ padding: "8px 10px", fontSize: 13, color: "#9ca3af" }}>ไม่พบลูกค้า</div>}
      {filtered.map((c, idx) => (
        <div
          key={c.id}
          data-idx={idx}
          onMouseDown={(e) => { e.preventDefault(); selectAt(idx); }}
          onMouseEnter={() => setHighlight(idx)}
          style={{
            padding: "8px 10px", fontSize: 13, cursor: "pointer",
            background: idx === highlight ? "#e0ddfb" : (c.id === value ? "#eeedfe" : "#fff"),
          }}
        >
          {display(c)}
        </div>
      ))}
    </div>
  );

  return (
    <div ref={wrapRef} style={{ position: "relative", minWidth }}>
      <input
        ref={inputRef}
        style={inputStyle}
        disabled={disabled}
        placeholder="ค้นหาลูกค้า..."
        value={open ? query : (selected ? display(selected) : "")}
        onFocus={() => { setOpen(true); setQuery(""); }}
        onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
        onKeyDown={handleKeyDown}
      />
      {dropdown}
    </div>
  );
}

// ===================================================================
// MAIN APP
// ===================================================================
export default function App() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [loginForm, setLoginForm] = useState({ username: "", password: "" });
  const [loginError, setLoginError] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [users, setUsers] = useState([
    { id: "U001", username: "changlo", password: "0260", name: "ผู้ดูแลระบบ", role: "admin" },
  ]);
  const [currentUser, setCurrentUser] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [printPreview, setPrintPreview] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);

  React.useEffect(() => {
    registerPrintPreview(setPrintPreview);
    registerConfirmDialog(setConfirmDialog);
  }, []);


  const handleLogin = () => {
    const user = users.find((u) => u.username === loginForm.username && u.password === loginForm.password);
    if (user) {
      setCurrentUser(user);
      setIsLoggedIn(true);
      setLoginError("");
    } else {
      setLoginError("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
    }
  };

  const handleLogout = () => {
    setIsLoggedIn(false);
    setCurrentUser(null);
    setLoginForm({ username: "", password: "" });
  };

  // หน้า Login


 const [products, setProducts] = useState(initialProducts);
  const [customers, setCustomers] = useState(initialCustomers);
  const [purchases, setPurchases] = useState(initialPurchases);
  const [sales, setSales] = useState(initialSales);
  const [storeBankAccounts, setStoreBankAccounts] = useState(initialStoreBankAccounts);


  // ตั้งค่ากิจการ (Company Settings) — ใช้ใน header ของใบรับ/ขายสินค้า
  // shopProfile — ชื่อ/โลโก้ใน sidebar (แยกจากข้อมูลบิล)
  const [shopProfile, setShopProfile] = useState({
    name: "วงจรกรีน",
    nameEn: "ระบบซื้อขายของเก่ารีไซเคิล",
    logo: "",   // base64
  });

  const [companySettings, setCompanySettings] = useState({
    name: "วงจรกรีน รีไซเคิล",
    nameEn: "",
    taxId: "",
    address: "",
    phone: "",
    email: "",
    website: "",
    logo: "",           // base64 image สำหรับบิล
    // ตั้งค่าเอกสาร
    purchaseTitle: "ใบรับสินค้า (รับซื้อของเก่า)",
    salesTitle: "ใบกำกับภาษี / Invoice",
    expenseVoucherTitle: "ใบสำคัญจ่าย",
    expenseVoucherNote: "",
    primaryColor: "#0f6e56",
    accentColor: "#185fa5",
    footerNote: "",
    showQrCode: false,
    showSignature: true,
    openingRevenue: 0,  // รายได้ยกมาก่อนเริ่มใช้แอพ
    openingCost: 0,     // ต้นทุนยกมาก่อนเริ่มใช้แอพ
    openingMonth: "",   // เดือนที่ยอดยกมามีผล (YYYY-MM)
  });

  const [withdrawals, setWithdrawals] = useState(initialWithdrawals);
  const [deposits, setDeposits] = useState(initialDeposits);
  const [prepayments, setPrepayments] = useState([]);
  const [deliveries, setDeliveries] = useState([]);
  const [bankTransfers, setBankTransfers] = useState([]);
  const [expenses, setExpenses] = useState(initialExpenses);
  const [loans, setLoans] = useState(initialLoans);
  const [assets, setAssets] = useState(initialAssets);
  const [shareholders, setShareholders] = useState(initialShareholders);
  const [dividendPayments, setDividendPayments] = useState([]);
  const [unitOptions, setUnitOptions] = useState(UNIT_OPTIONS_DEFAULT);
  const [expenseCategories, setExpenseCategories] = useState(EXPENSE_SUBCATEGORIES_DEFAULT);
  const [productCategories, setProductCategories] = useState(PRODUCT_TYPES);

  const inventory = useMemo(() => computeInventory(products, purchases, sales, withdrawals), [products, purchases, sales, withdrawals]);

  // ===== Supabase Sync =====
  const [dbLoaded, setDbLoaded] = useState(false)
  const syncStatus = useSyncStatus()
  useProductsRealtime(setProducts, dbLoaded)

  // โหลดข้อมูลจาก Supabase ครั้งแรก
  useEffect(() => {
    if (!isSupabaseReady) { setDbLoaded(true);
; return }
    // ลบรายการที่ id ซ้ำออกก่อน set state (ป้องกันบิลซ้ำจากปัญหา sync เก่า)
    const dedup = (arr) => {
      if (!Array.isArray(arr)) return arr
      const seen = new Set()
      return arr.filter(item => {
        if (!item || !item.id) return true
        if (seen.has(item.id)) return false
        seen.add(item.id)
        return true
      })
    }

    loadAllFromSupabase().then(data => {
      if (data) {
        if (data.customers)     setCustomers(dedup(data.customers))
        if (data.purchases)     setPurchases(dedup(data.purchases))
        if (data.sales)         setSales(dedup(data.sales))
        if (data.withdrawals)   setWithdrawals(dedup(data.withdrawals))
        if (data.deposits)      setDeposits(dedup(data.deposits))
        if (data.bankTransfers) setBankTransfers(dedup(data.bankTransfers))
        if (data.expenses)      setExpenses(dedup(data.expenses))
        if (data.loans)         setLoans(dedup(data.loans))
        if (data.storeBankAccounts) setStoreBankAccounts(dedup(data.storeBankAccounts))
        if (data.shopProfile)   setShopProfile(data.shopProfile)
        if (data.companySettings) setCompanySettings(data.companySettings)
        if (data.users)         setUsers(data.users)
        if (data.unitOptions)   setUnitOptions(data.unitOptions)
        if (data.expenseCategories) setExpenseCategories(data.expenseCategories)
        if (data.productCategories) setProductCategories(data.productCategories)
        if (data.assets) setAssets(dedup(data.assets))
        if (data.shareholders) setShareholders(data.shareholders)
        if (data.dividendPayments) setDividendPayments(dedup(data.dividendPayments))
        if (data.deliveries) setDeliveries(dedup(data.deliveries))
        if (data.prepayments) setPrepayments(dedup(data.prepayments))
      }
      setDbLoaded(true)
    })
  }, [])

  // ปุ่ม "โหลดข้อมูลล่าสุด" — โหลดจาก Supabase ใหม่ทั้งหมดทันที
  const [isReloading, setIsReloading] = useState(false)
  const reloadFromSupabase = async () => {
    if (!isSupabaseReady || isReloading) return
    setIsReloading(true)
    const data = await loadAllFromSupabase()
    if (data) {
      if (data.customers)       setCustomers(dedup(data.customers))
      if (data.purchases)       setPurchases(dedup(data.purchases))
      if (data.sales)           setSales(dedup(data.sales))
      if (data.withdrawals)     setWithdrawals(dedup(data.withdrawals))
      if (data.deposits)        setDeposits(dedup(data.deposits))
      if (data.bankTransfers)   setBankTransfers(dedup(data.bankTransfers))
      if (data.expenses)        setExpenses(dedup(data.expenses))
      if (data.loans)           setLoans(dedup(data.loans))
      if (data.storeBankAccounts) setStoreBankAccounts(dedup(data.storeBankAccounts))
      if (data.shopProfile)     setShopProfile(data.shopProfile)
      if (data.companySettings) setCompanySettings(data.companySettings)
      if (data.users)           setUsers(data.users)
      if (data.unitOptions)     setUnitOptions(data.unitOptions)
      if (data.expenseCategories) setExpenseCategories(data.expenseCategories)
      if (data.productCategories) setProductCategories(data.productCategories)
      if (data.assets)          setAssets(dedup(data.assets))
      if (data.shareholders)    setShareholders(data.shareholders)
      if (data.dividendPayments) setDividendPayments(dedup(data.dividendPayments))
      if (data.deliveries) setDeliveries(dedup(data.deliveries))
      if (data.prepayments) setPrepayments(dedup(data.prepayments))
    }
    setIsReloading(false)
  }

  // Auto-sync แต่ละ state ไปยัง Supabase
  useSupabaseSync('customers',         customers,         setCustomers,         dbLoaded)
  useSupabaseSync('purchases',         purchases,         setPurchases,         dbLoaded)
  useSupabaseSync('sales',             sales,             setSales,             dbLoaded)
  useSupabaseSync('withdrawals',       withdrawals,       setWithdrawals,       dbLoaded)
  useSupabaseSync('deposits',          deposits,          setDeposits,          dbLoaded)
  useSupabaseSync('bankTransfers',     bankTransfers,     setBankTransfers,     dbLoaded)
  useSupabaseSync('expenses',          expenses,          setExpenses,          dbLoaded)
  useSupabaseSync('loans',             loans,             setLoans,             dbLoaded)
  useSupabaseSync('storeBankAccounts', storeBankAccounts, setStoreBankAccounts, dbLoaded)
  useSupabaseSync('shopProfile',       shopProfile,       setShopProfile,       dbLoaded)
  useSupabaseSync('companySettings',   companySettings,   setCompanySettings,   dbLoaded)
  useSupabaseSync('unitOptions',       unitOptions,       setUnitOptions,       dbLoaded)
  useSupabaseSync('expenseCategories', expenseCategories, setExpenseCategories, dbLoaded)
  useSupabaseSync('productCategories', productCategories, setProductCategories, dbLoaded)
  useSupabaseSync('assets',            assets,            setAssets,            dbLoaded)
  useSupabaseSync('shareholders',      shareholders,      setShareholders,      dbLoaded)
  useSupabaseSync('dividendPayments',  dividendPayments,  setDividendPayments,  dbLoaded)
  useSupabaseSync('deliveries',        deliveries,        setDeliveries,        dbLoaded)
  useSupabaseSync('prepayments',       prepayments,       setPrepayments,       dbLoaded)

useEffect(() => {
  loadProducts().then(setProducts);
}, []);

  const navItems = [
    { key: "dashboard",         label: "แดชบอร์ด",              icon: LayoutDashboard },
    { key: "purchases",         label: "ใบรับสินค้า",            icon: ArrowDownToLine },
    { key: "withdrawals",       label: "ใบเบิกสินค้า",           icon: PackageMinus },
    { key: "sales",             label: "ขายสินค้า",              icon: ShoppingCart },
    { key: "expenses",          label: "ค่าใช้จ่าย",             icon: Receipt },
    { key: "payments",          label: "รับชำระ/จ่ายชำระ",       icon: BadgeDollarSign },
    { key: "deposits",          label: "เงินมัดจำ",              icon: Wallet },
    { key: "prepayments",       label: "รับล่วงหน้า",             icon: BadgeDollarSign },
    { key: "banktransfer",      label: "โยกเงินระหว่างธนาคาร",   icon: ArrowLeftRight },
    { key: "delivery",          label: "ใบส่งสินค้า",            icon: Truck },
    { key: "inventory",         label: "สต๊อกสินค้า",            icon: Boxes },
    { key: "bankaccounts",      label: "บัญชีธนาคารร้าน",        icon: Landmark },
    { key: "loans",             label: "เงินกู้ยืม/เช่าซื้อ",    icon: Banknote },
    { key: "assets",            label: "ทะเบียนทรัพย์สิน",       icon: Building2 },
    { key: "report",            label: "รายงานกำไร",             icon: BarChart3 },
    { key: "tax",               label: "สรุปภาษี",               icon: PieChart },
    { key: "products",          label: "ข้อมูลสินค้า",            icon: Package },
    { key: "customers",         label: "ข้อมูลลูกค้า",            icon: Users },
    { key: "expenseCategories", label: "หมวดหมู่ค่าใช้จ่าย",     icon: Tag },
    { key: "settings",          label: "ตั้งค่ากิจการ",           icon: Settings },
  ];


  if (!isLoggedIn) {
    return (
      <div style={{ minHeight: "100vh", background: "linear-gradient(135deg, #0c443c 0%, #1d9e75 100%)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Noto Sans Thai', 'Inter', system-ui, sans-serif" }}>
        <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;500;600;700&display=swap" />
        <div style={{ background: "#fff", borderRadius: 20, padding: "40px 36px", width: "100%", maxWidth: 400, boxShadow: "0 20px 60px rgba(0,0,0,0.2)" }}>
          {/* Logo / ชื่อแอพ */}
          <div style={{ textAlign: "center", marginBottom: 32 }}>
            {shopProfile.logo ? (
              <img src={shopProfile.logo} alt="logo"
                style={{ width: 72, height: 72, objectFit: "contain", borderRadius: 16, background: "#f3f4f6", padding: 8, margin: "0 auto 16px", display: "block" }} />
            ) : (
              <div style={{ width: 64, height: 64, background: "linear-gradient(135deg, #0c443c, #1d9e75)", borderRadius: 16, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px" }}>
                <Boxes size={32} color="#fff" />
              </div>
            )}
            <div style={{ fontWeight: 700, fontSize: 22, color: "#0c443c" }}>{shopProfile.name || "วงจรกรีน รีไซเคิล"}</div>
            <div style={{ fontSize: 13, color: "#9ca3af", marginTop: 4 }}>{shopProfile.nameEn || "ระบบซื้อขายของเก่ารีไซเคิล"}</div>
          </div>

          {/* Form */}
          <div style={{ marginBottom: 16 }}>
            <label style={{ display: "block", fontSize: 13, fontWeight: 600, color: "#374151", marginBottom: 6 }}>ชื่อผู้ใช้</label>
            <input
              style={{ width: "100%", padding: "10px 14px", border: "1.5px solid #d1d5db", borderRadius: 10, fontSize: 14, outline: "none", boxSizing: "border-box", fontFamily: "inherit" }}
              placeholder="username"
              value={loginForm.username}
              onChange={(e) => setLoginForm({ ...loginForm, username: e.target.value })}
              onKeyDown={(e) => e.key === "Enter" && handleLogin()}
              autoFocus
            />
          </div>
          <div style={{ marginBottom: 20 }}>
            <label style={{ display: "block", fontSize: 13, fontWeight: 600, color: "#374151", marginBottom: 6 }}>รหัสผ่าน</label>
            <div style={{ position: "relative" }}>
              <input
                type={showPassword ? "text" : "password"}
                style={{ width: "100%", padding: "10px 40px 10px 14px", border: "1.5px solid #d1d5db", borderRadius: 10, fontSize: 14, outline: "none", boxSizing: "border-box", fontFamily: "inherit" }}
                placeholder="password"
                value={loginForm.password}
                onChange={(e) => setLoginForm({ ...loginForm, password: e.target.value })}
                onKeyDown={(e) => e.key === "Enter" && handleLogin()}
              />
              <button
                onClick={() => setShowPassword(!showPassword)}
                style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: "#9ca3af", padding: 0 }}
              >
                {showPassword ? <X size={16} /> : <ArrowRight size={16} />}
              </button>
            </div>
          </div>

          {loginError && (
            <div style={{ background: "#fcebeb", border: "1px solid #f5c2c2", borderRadius: 8, padding: "8px 12px", fontSize: 13, color: "#a32d2d", marginBottom: 14, textAlign: "center" }}>
              {loginError}
            </div>
          )}

          <button
            onClick={handleLogin}
            style={{ width: "100%", padding: "12px", background: "linear-gradient(135deg, #0c443c, #1d9e75)", border: "none", borderRadius: 10, color: "#fff", fontWeight: 700, fontSize: 15, cursor: "pointer", fontFamily: "inherit" }}
          >
            เข้าสู่ระบบ
          </button>

          <div style={{ marginTop: 16, fontSize: 12, color: "#9ca3af", textAlign: "center" }}>
            ค่าเริ่มต้น: admin / 1234
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", minHeight: "100vh", fontFamily: "'Noto Sans Thai', 'Inter', system-ui, sans-serif", background: "#f3f4f1", color: "#1f2937" }}>
      <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;500;600;700&family=JetBrains+Mono:wght@500&display=swap" />

      <PrintPreviewOverlay preview={printPreview} onClose={() => setPrintPreview(null)} />
      <ConfirmDialog pending={confirmDialog} onClose={() => setConfirmDialog(null)} />

      {/* Sidebar — fixed, independent scroll */}
      <div style={{ width: sidebarOpen ? 220 : 64, background: "#0c443c", color: "#e1f5ee", display: "flex", flexDirection: "column", flexShrink: 0, transition: "width 0.2s ease", height: "100vh", overflowY: "auto", overflowX: "auto", position: "fixed", top: 0, left: 0, zIndex: 10 }}>
        <div style={{ padding: sidebarOpen ? "16px 18px" : "16px 10px", borderBottom: "1px solid rgba(255,255,255,0.08)", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          {sidebarOpen ? (
            <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0, flex: 1 }}>
              {/* โลโก้ หรือ icon สำรอง */}
              {shopProfile.logo ? (
                <img
                  src={shopProfile.logo}
                  alt="logo"
                  style={{ width: 38, height: 38, borderRadius: 8, objectFit: "contain", background: "#fff", padding: 3, flexShrink: 0 }}
                />
              ) : (
                <div style={{ width: 38, height: 38, borderRadius: 8, background: "#1d9e75", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <Boxes size={20} color="#04342c" />
                </div>
              )}
              <div style={{ minWidth: 0, overflow: "hidden" }}>
                <div style={{ fontWeight: 700, fontSize: 14, lineHeight: 1.2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                  {shopProfile.name || "วงจรกรีน"}
                </div>
                <div style={{ fontSize: 10, color: "#9fe1cb", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                  {shopProfile.nameEn || "ระบบซื้อขายของเก่ารีไซเคิล"}
                </div>
              </div>
            </div>
          ) : (
            /* ย่อเมนู: แสดงแค่โลโก้หรือ icon */
            shopProfile.logo ? (
              <img
                src={shopProfile.logo}
                alt="logo"
                style={{ width: 38, height: 38, borderRadius: 8, objectFit: "contain", background: "#fff", padding: 3, margin: "0 auto", display: "block" }}
              />
            ) : (
              <div style={{ width: 38, height: 38, borderRadius: 8, background: "#1d9e75", display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto" }}>
                <Boxes size={20} color="#04342c" />
              </div>
            )
          )}
          {sidebarOpen && (
            <button
              onClick={() => setSidebarOpen(false)}
              title="ย่อเมนู"
              style={{ background: "rgba(255,255,255,0.08)", border: "none", borderRadius: 6, width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center", color: "#d1ede1", cursor: "pointer", flexShrink: 0 }}
            >
              <ChevronLeft size={16} />
            </button>
          )}
        </div>
        {!sidebarOpen && (
          <div style={{ display: "flex", justifyContent: "center", padding: "8px 0", borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
            <button
              onClick={() => setSidebarOpen(true)}
              title="ขยายเมนู"
              style={{ background: "rgba(255,255,255,0.08)", border: "none", borderRadius: 6, width: 28, height: 28, display: "flex", alignItems: "center", justifyContent: "center", color: "#d1ede1", cursor: "pointer" }}
            >
              <Menu size={16} />
            </button>
          </div>
        )}
        <nav style={{ flex: 1, padding: sidebarOpen ? "12px 10px" : "12px 8px" }}>
          {navItems.map((item) => {
            const Icon = item.icon;
            const active = tab === item.key;
            return (
              <button
                key={item.key}
                onClick={() => setTab(item.key)}
                title={!sidebarOpen ? item.label : undefined}
                style={{
                  width: "100%", display: "flex", alignItems: "center", gap: 10,
                  justifyContent: sidebarOpen ? "flex-start" : "center",
                  padding: sidebarOpen ? "10px 12px" : "10px 0", marginBottom: 4, borderRadius: 8, border: "none",
                  background: active ? "#1d9e75" : "transparent",
                  color: active ? "#04342c" : "#d1ede1",
                  fontWeight: active ? 600 : 500, fontSize: 14, cursor: "pointer", textAlign: "left",
                  transition: "background 0.15s", overflow: "hidden", whiteSpace: "nowrap",
                }}
              >
                <Icon size={18} style={{ flexShrink: 0 }} />
                {sidebarOpen && item.label}
              </button>
            );
          })}
        </nav>

        {/* ผู้ใช้งาน + ออกจากระบบ */}
        {/* ปุ่มโหลดข้อมูลล่าสุด + sync status */}
        {isSupabaseReady && (
          <div style={{ padding: sidebarOpen ? "8px 12px" : "8px 6px" }}>
            {sidebarOpen && (
              <div style={{
                display: "flex", alignItems: "center", gap: 6,
                padding: "5px 10px", borderRadius: 6, marginBottom: 6,
                background: syncStatus === 'saving' ? "rgba(251,191,36,0.15)"
                  : syncStatus === 'error' ? "rgba(239,68,68,0.15)"
                  : "rgba(29,158,117,0.1)",
                border: syncStatus === 'saving' ? "1px solid rgba(251,191,36,0.4)"
                  : syncStatus === 'error' ? "1px solid rgba(239,68,68,0.4)"
                  : "1px solid rgba(29,158,117,0.2)",
              }}>
                <div style={{
                  width: 8, height: 8, borderRadius: "50%", flexShrink: 0,
                  background: syncStatus === 'saving' ? "#fbbf24"
                    : syncStatus === 'error' ? "#ef4444"
                    : "#1d9e75",
                  animation: syncStatus === 'saving' ? "pulse 1s ease-in-out infinite" : "none",
                }} />
                <span style={{
                  fontSize: 11, fontWeight: 600,
                  color: syncStatus === 'saving' ? "#fbbf24"
                    : syncStatus === 'error' ? "#fca5a5"
                    : "#9fe1cb",
                }}>
                  {syncStatus === 'saving' ? "กำลังบันทึก..." : syncStatus === 'error' ? "บันทึกไม่สำเร็จ!" : "บันทึกแล้ว ✓"}
                </span>
              </div>
            )}
            <button
              onClick={reloadFromSupabase}
              disabled={isReloading}
              title="โหลดข้อมูลล่าสุดจาก Supabase"
              style={{
                width: "100%", display: "flex", alignItems: "center", justifyContent: sidebarOpen ? "flex-start" : "center",
                gap: 8, padding: sidebarOpen ? "8px 12px" : "8px 6px", borderRadius: 8,
                background: isReloading ? "rgba(255,255,255,0.05)" : "rgba(29,158,117,0.2)",
                border: "1px solid rgba(29,158,117,0.4)", color: "#9fe1cb", cursor: isReloading ? "not-allowed" : "pointer",
                fontSize: 12, fontWeight: 600, transition: "all 0.15s",
              }}
            >
              <Download size={15} style={{ animation: isReloading ? "spin 1s linear infinite" : "none", flexShrink: 0 }} />
              {sidebarOpen && (isReloading ? "กำลังโหลด..." : "โหลดข้อมูลล่าสุด")}
            </button>
            <style>{`
              @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
              @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.4; } }
            `}</style>
          </div>
        )}
        <div style={{ borderTop: "1px solid rgba(255,255,255,0.08)", padding: sidebarOpen ? "12px 16px" : "12px 8px" }}>
          {sidebarOpen ? (
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                <div style={{ width: 30, height: 30, borderRadius: "50%", background: "#1d9e75", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <Users size={15} color="#04342c" />
                </div>
                <div>
                  <div style={{ fontSize: 12, fontWeight: 600, color: "#e1f5ee" }}>{currentUser?.name}</div>
                  <div style={{ fontSize: 10, color: "#9fe1cb" }}>{currentUser?.role === "admin" ? "ผู้ดูแลระบบ" : "ผู้ใช้งาน"}</div>
                </div>
              </div>
              <button
                onClick={handleLogout}
                style={{ width: "100%", padding: "7px 10px", background: "rgba(255,80,80,0.15)", border: "1px solid rgba(255,80,80,0.3)", borderRadius: 8, color: "#fca5a5", fontSize: 12, fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 6, fontFamily: "inherit" }}
              >
                <X size={13} /> ออกจากระบบ
              </button>
            </div>
          ) : (
            <button
              onClick={handleLogout}
              title="ออกจากระบบ"
              style={{ width: "100%", padding: "8px", background: "rgba(255,80,80,0.15)", border: "1px solid rgba(255,80,80,0.3)", borderRadius: 8, color: "#fca5a5", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}
            >
              <X size={16} />
            </button>
          )}
        </div>
      </div>

      {/* Main content — independently scrollable */}
      <div style={{ flex: 1, padding: "28px 32px", overflowY: "auto", overflowX: "auto", minHeight: "100vh", marginLeft: sidebarOpen ? 220 : 64, transition: "margin-left 0.2s ease", boxSizing: "border-box", width: sidebarOpen ? "calc(100vw - 220px)" : "calc(100vw - 64px)" }}>        {tab === "dashboard" && <Dashboard products={products} customers={customers} purchases={purchases} sales={sales} inventory={inventory} expenses={expenses} loans={loans} storeBankAccounts={storeBankAccounts} deposits={deposits} bankTransfers={bankTransfers} expenseCategories={expenseCategories} prepayments={prepayments} />}
        {tab === "products" && <ProductsTab products={products} setProducts={setProducts} unitOptions={unitOptions} setUnitOptions={setUnitOptions} productCategories={productCategories} setProductCategories={setProductCategories} />}
        {tab === "customers" && <CustomersTab customers={customers} setCustomers={setCustomers} />}
        {tab === "purchases" && <PurchasesTab products={products} customers={customers} purchases={purchases} setPurchases={setPurchases} storeBankAccounts={storeBankAccounts} deposits={deposits} companySettings={companySettings} />}
        {tab === "withdrawals" && <WithdrawalsTab products={products} purchases={purchases} sales={sales} setSales={setSales} withdrawals={withdrawals} setWithdrawals={setWithdrawals} inventory={inventory} customers={customers} companySettings={companySettings} />}
        {tab === "sales" && <SalesTab products={products} customers={customers} sales={sales} setSales={setSales} inventory={inventory} withdrawals={withdrawals} storeBankAccounts={storeBankAccounts} companySettings={companySettings} />}
        {tab === "payments" && <PaymentsTab purchases={purchases} setPurchases={setPurchases} sales={sales} setSales={setSales} customers={customers} storeBankAccounts={storeBankAccounts} deposits={deposits} expenses={expenses} setExpenses={setExpenses} companySettings={companySettings} setCompanySettings={setCompanySettings} bankTransfers={bankTransfers} />}
        {tab === "delivery" && <DeliveryTab deliveries={deliveries} setDeliveries={setDeliveries} products={products} customers={customers} sales={sales} companySettings={companySettings} />}
        {tab === "inventory" && <InventoryTab products={products} inventory={inventory} storeBankAccounts={storeBankAccounts} />}
        {tab === "deposits" && <DepositsTab customers={customers} setCustomers={setCustomers} deposits={deposits} setDeposits={setDeposits} purchases={purchases} storeBankAccounts={storeBankAccounts} />}
        {tab === "prepayments" && <PrepaymentsTab customers={customers} setCustomers={setCustomers} prepayments={prepayments} setPrepayments={setPrepayments} sales={sales} storeBankAccounts={storeBankAccounts} />}
        {tab === "expenses" && <ExpensesTab expenses={expenses} setExpenses={setExpenses} storeBankAccounts={storeBankAccounts} loans={loans} setLoans={setLoans} expenseCategories={expenseCategories} setExpenseCategories={setExpenseCategories} companySettings={companySettings} customers={customers} />}
        {tab === "expenseCategories" && <ExpenseCategoriesTab expenseCategories={expenseCategories} setExpenseCategories={setExpenseCategories} expenses={expenses} setExpenses={setExpenses} />}
        {tab === "loans" && <LoansTab loans={loans} setLoans={setLoans} expenses={expenses} customers={customers} />}
        {tab === "bankaccounts" && <StoreBankAccountsTab accounts={storeBankAccounts} setAccounts={setStoreBankAccounts} purchases={purchases} sales={sales} expenses={expenses} deposits={deposits} bankTransfers={bankTransfers} customers={customers} />}
        {tab === "banktransfer" && <BankTransferTab storeBankAccounts={storeBankAccounts} bankTransfers={bankTransfers} setBankTransfers={setBankTransfers} />}
        {tab === "assets" && <AssetsTab assets={assets} setAssets={setAssets} />}
        {tab === "settings" && <CompanySettingsTab settings={companySettings} setSettings={setCompanySettings} shopProfile={shopProfile} setShopProfile={setShopProfile} />}
        {tab === "report" && <MonthlyReportTab purchases={purchases} sales={sales} expenses={expenses} inventory={inventory} withdrawals={withdrawals} expenseCategories={expenseCategories} shareholders={shareholders} setShareholders={setShareholders} dividendPayments={dividendPayments} setDividendPayments={setDividendPayments} companySettings={companySettings} setCompanySettings={setCompanySettings} />}
        {tab === "tax" && <TaxSummaryTab purchases={purchases} sales={sales} expenses={expenses} />}
      </div>

    </div>
  );
}

// ===================================================================
// DASHBOARD
// ===================================================================
function Dashboard({ products, customers, purchases, sales, inventory, expenses, loans, storeBankAccounts, deposits, bankTransfers, expenseCategories, prepayments }) {
  // ---------- หมวดหมู่แดชบอร์ด ----------
  const [dashSubTab, setDashSubTab] = useState("purchases"); // "purchases" | "sales" | "expenses" | "stock" | "loans"
  const [expandedStockTypes, setExpandedStockTypes] = useState({}); // { [type]: bool } ติ๊กเลือกเพื่อดูรายการสินค้าในประเภทนั้น

  // ---------- ตัวเลือกช่วงเวลา: รายวัน / ช่วงวันที่ (เลือกเอง) / ทั้งหมด ----------
  const today = new Date().toISOString().slice(0, 10);
  const [periodMode, setPeriodMode] = useState("day"); // "all" | "day" | "range"
  const [periodDate, setPeriodDate] = useState(today);
  const [rangeStart, setRangeStart] = useState(today);
  const [rangeEnd, setRangeEnd] = useState(today);

  // ช่วงวันที่ตามตัวเลือก: คืนค่า {start, end} (รวมทั้งสองค่า) หรือ null = ไม่จำกัด (ทั้งหมด)
  const dateRange = useMemo(() => {
    if (periodMode === "day") return { start: periodDate, end: periodDate };
    if (periodMode === "range") {
      const start = rangeStart <= rangeEnd ? rangeStart : rangeEnd;
      const end = rangeStart <= rangeEnd ? rangeEnd : rangeStart;
      return { start, end };
    }
    return null;
  }, [periodMode, periodDate, rangeStart, rangeEnd]);

  const inRange = (dateStr) => {
    if (!dateRange) return true;
    return dateStr >= dateRange.start && dateStr <= dateRange.end;
  };

  // ---------- กรองรายการซื้อ/ขายตามช่วงเวลา ----------
  const filteredPurchases = purchases.filter((po) => po.status === "อนุมัติแล้ว" && inRange(po.date));
  const filteredSales = sales.filter((inv) => inRange(inv.date));

  // มูลค่าซื้อ ก่อน VAT (ต้นทุนสินค้าที่ใช้คำนวณสต๊อก/กำไร)
  const totalPurchaseValue = filteredPurchases.reduce((sum, po) => sum + po.items.reduce((s, it) => s + (it.net || 0) * (it.price || 0), 0), 0);
  const totalSalesValue = filteredSales.reduce((sum, inv) => {
    const subtotal = inv.items.reduce((s, it) => s + it.net * it.price, 0);
    const afterDiscount = subtotal - (inv.discount || 0);
    const vat = afterDiscount * ((inv.vatRate || 0) / 100);
    return sum + afterDiscount + vat;
  }, 0);

  // ---------- ยอดยกมาของหมวดหมู่ย่อยค่าใช้จ่าย (นับรวมเฉพาะตอนช่วงเวลาที่เลือกตรงกับเดือนที่ระบุไว้ หรือเลือก "ทั้งหมด") ----------
  const openingBalanceApplies = (openingMonth) => {
    if (!openingMonth) return false;
    if (!dateRange) return true; // periodMode = "all" → นับรวมยอดยกมาทุกตัว
    // periodMode = "day" หรือ "range" → นับรวมเฉพาะถ้าช่วงที่เลือกอยู่ในเดือนที่ระบุไว้เท่านั้น
    const monthStart = `${openingMonth}-01`;
    const lastDay = new Date(Number(openingMonth.slice(0,4)), Number(openingMonth.slice(5,7)), 0).getDate();
    const monthEnd = `${openingMonth}-${String(lastDay).padStart(2,"0")}`;
    return dateRange.start >= monthStart && dateRange.end <= monthEnd;
  };

  const expenseOpeningRows = useMemo(() => {
    const rows = [];
    Object.entries(expenseCategories || {}).forEach(([main, subs]) => {
      if (main !== "ค่าใช้จ่าย") return;
      (subs || []).forEach((s) => {
        if (typeof s === "string") return; // ของเดิมไม่มียอดยกมา
        if (Number(s.openingBalance) > 0 && openingBalanceApplies(s.openingMonth)) {
          rows.push({ subCategory: s.name, amount: Number(s.openingBalance) });
        }
      });
    });
    return rows;
  }, [expenseCategories, dateRange]);

  const totalExpensesOpening = expenseOpeningRows.reduce((s, r) => s + r.amount, 0);

  // รวมรายได้ = ยอดขาย + รายได้อื่นยกมา
  // รวมค่าใช้จ่าย = ค่าใช้จ่ายจริง + ค่าใช้จ่ายยกมา
  const totalExpenses = (expenses || []).filter((e) => inRange(e.billDate || e.date)).reduce((s, e) => {
    const items = (e.items && e.items.length > 0) ? e.items : [{ mainCategory: e.mainCategory || e.category, amount: e.amount }];
    return s + items.filter((it) => it.mainCategory === "ค่าใช้จ่าย").reduce((s2, it) => s2 + (Number(it.amount) || 0), 0);
  }, 0) + totalExpensesOpening;

  // ---------- ค่าใช้จ่ายแบ่งตามหมวดหมู่ย่อย ----------
  const expensesBySubCategory = useMemo(() => {
    const groups = {};
    // จากค่าใช้จ่ายจริง
    (expenses || []).filter((e) => inRange(e.billDate || e.date)).forEach((e) => {
      const items = (e.items && e.items.length > 0) ? e.items : [{ mainCategory: e.mainCategory || e.category, subCategory: e.subCategory, amount: e.amount }];
      items.filter((it) => it.mainCategory === "ค่าใช้จ่าย").forEach((it) => {
        const sub = it.subCategory || "อื่นๆ";
        if (!groups[sub]) groups[sub] = { subCategory: sub, amount: 0, count: 0 };
        groups[sub].amount += Number(it.amount) || 0;
        groups[sub].count += 1;
      });
    });
    // จากยอดยกมา
    expenseOpeningRows.forEach((r) => {
      if (!groups[r.subCategory]) groups[r.subCategory] = { subCategory: r.subCategory, amount: 0, count: 0 };
      groups[r.subCategory].amount += r.amount;
    });
    return Object.values(groups).sort((a, b) => b.amount - a.amount);
  }, [expenses, dateRange, expenseOpeningRows]);


  const totalStockValue = inventory.summary.reduce((s, x) => s + x.totalCost, 0);

  // ===== ยอดยกมา =====
  const totalOpeningStockQty   = products.reduce((s, p) => s + (Number(p.openingQty)  || 0), 0);
  const totalOpeningStockValue = products.reduce((s, p) => s + (Number(p.openingQty) || 0) * (Number(p.openingCost) || 0), 0);

  // ยอดยกมาธนาคาร (จาก storeBankAccounts.openingBalance)
  const totalOpeningBankBalance = storeBankAccounts.reduce((s, a) => s + (Number(a.openingBalance) || 0), 0);

  // มีข้อมูลยกมาไหม
  const hasOpeningData = totalOpeningStockValue > 0 || totalOpeningBankBalance > 0;

  // ---------- สต๊อกคงเหลือ แบ่งตามประเภทสินค้า (สำหรับตัวกรองดรอปดาวน์) ----------
  const stockByType = useMemo(() => {
    const groups = {}; // type -> { type, qty, value, items: [...] }
    inventory.summary.forEach((s) => {
      const p = products.find((pr) => pr.id === s.productId);
      const type = p?.type || "ไม่ระบุประเภท";
      if (!groups[type]) groups[type] = { type, qty: 0, value: 0, items: [] };
      groups[type].qty += s.qty;
      groups[type].value += s.totalCost;
      groups[type].items.push(s);
    });
    return Object.values(groups)
      .map((g) => ({ ...g, avgCost: g.qty > 0 ? g.value / g.qty : 0 }))
      .sort((a, b) => b.value - a.value);
  }, [inventory.summary, products]);

  // ---------- คงเหลือสินเชื่อ/เงินกู้ — ยอดเงินต้นคงเหลือรวมทุกสัญญา ณ ปัจจุบัน ----------
  const totalLoanRemaining = (loans || []).reduce((sum, loan) => {
    const schedule = computeAmortizationSchedule(loan);
    const paidCount = (loan.paidInstallments || []).length;
    const nextInstallment = schedule.find((s) => s.no === paidCount + 1);
    return sum + (nextInstallment ? nextInstallment.remainingBalance + nextInstallment.principalPortion : 0);
  }, 0);

  // ---------- ยอดคงเหลือแบงค์ — สุทธิเงินที่จ่ายออกจากบัญชีร้านแต่ละบัญชี (สะสมทั้งหมด ไม่ขึ้นกับช่วงเวลา) ----------
  const bankOutflows = useMemo(() => {
    const out = {}; // bankId -> total
    const add = (bankId, amount) => {
      if (!bankId || bankId === "CASH" || bankId === "DEPOSIT" || bankId === "PREPAYMENT") return;
      out[bankId] = (out[bankId] || 0) + amount;
    };
    purchases.forEach((po) => (po.payments || []).forEach((p) => add(p.fromStoreBankId, Number(p.amount) || 0)));
    (deposits || []).forEach((d) => add(d.fromStoreBankId, Number(d.amount) || 0));
    (expenses || []).forEach((e) => (e.payments || []).forEach((p) => add(p.fromStoreBankId, Number(p.amount) || 0)));
    (bankTransfers || []).forEach((t) => add(t.fromBankId, Number(t.amount) || 0));
    return out;
  }, [purchases, deposits, expenses, bankTransfers]);

  // ---------- ยอดรับเข้าแบงค์ — เงินที่รับเข้าบัญชีร้านแต่ละบัญชีจากการขาย (สะสมทั้งหมด ไม่ขึ้นกับช่วงเวลา) ----------
  const bankInflows = useMemo(() => {
    const inn = {}; // bankId -> total
    const add = (bankId, amount) => {
      if (!bankId || bankId === "CASH" || bankId === "DEPOSIT" || bankId === "PREPAYMENT") return;
      inn[bankId] = (inn[bankId] || 0) + amount;
    };
    sales.forEach((inv) => (inv.payments || []).forEach((p) => add(p.toStoreBankId, Number(p.amount) || 0)));
    (prepayments || []).forEach((p) => add(p.toStoreBankId, Number(p.amount) || 0));
    (bankTransfers || []).forEach((t) => add(t.toBankId, Number(t.amount) || 0));
    return inn;
  }, [sales, bankTransfers, prepayments]);


  // ---------- ซื้อ/ขาย แบ่งตามประเภทสินค้า และแบ่งตามรายการสินค้า ----------
  const prodInfo = (id) => products.find((p) => p.id === id);

  const purchaseByType = useMemo(() => {
    const groups = {};
    filteredPurchases.forEach((po) => {
      po.items.forEach((it) => {
        const p = prodInfo(it.productId);
        const type = p?.type || "ไม่ระบุประเภท";
        const value = (it.net || 0) * (it.price || 0);
        const qty = it.net || 0;
        if (!groups[type]) groups[type] = { type, qty: 0, value: 0 };
        groups[type].qty += qty;
        groups[type].value += value;
      });
    });
    return Object.values(groups).map((g) => ({ ...g, avgCost: g.qty > 0 ? g.value / g.qty : 0 })).sort((a, b) => b.value - a.value);
  }, [filteredPurchases, products]);

  const purchaseByProduct = useMemo(() => {
    const groups = {};
    filteredPurchases.forEach((po) => {
      po.items.forEach((it) => {
        const value = (it.net || 0) * (it.price || 0);
        const qty = it.net || 0;
        if (!groups[it.productId]) groups[it.productId] = { productId: it.productId, qty: 0, value: 0 };
        groups[it.productId].qty += qty;
        groups[it.productId].value += value;
      });
    });
    return Object.values(groups).map((g) => ({ ...g, avgCost: g.qty > 0 ? g.value / g.qty : 0 })).sort((a, b) => b.value - a.value);
  }, [filteredPurchases]);

  const salesByType = useMemo(() => {
    const groups = {};
    filteredSales.forEach((inv) => {
      inv.items.forEach((it) => {
        const p = prodInfo(it.productId);
        const type = p?.type || "ไม่ระบุประเภท";
        const value = (it.net || 0) * (it.price || 0);
        const qty = it.net || 0;
        if (!groups[type]) groups[type] = { type, qty: 0, value: 0 };
        groups[type].qty += qty;
        groups[type].value += value;
      });
    });
    return Object.values(groups).map((g) => ({ ...g, avgCost: g.qty > 0 ? g.value / g.qty : 0 })).sort((a, b) => b.value - a.value);
  }, [filteredSales, products]);

  const salesByProduct = useMemo(() => {
    const groups = {};
    filteredSales.forEach((inv) => {
      inv.items.forEach((it) => {
        const value = (it.net || 0) * (it.price || 0);
        const qty = it.net || 0;
        if (!groups[it.productId]) groups[it.productId] = { productId: it.productId, qty: 0, value: 0 };
        groups[it.productId].qty += qty;
        groups[it.productId].value += value;
      });
    });
    return Object.values(groups).map((g) => ({ ...g, avgCost: g.qty > 0 ? g.value / g.qty : 0 })).sort((a, b) => b.value - a.value);
  }, [filteredSales]);

  const prodName = (id) => products.find((p) => p.id === id)?.name || id;
  const prodUnit = (id) => products.find((p) => p.id === id)?.unit || "";

  const purchaseCard = { label: "มูลค่าซื้อ ก่อน VAT (อนุมัติแล้ว)", value: fmt(totalPurchaseValue), suffix: "บาท", icon: ArrowDownToLine, color: "#d85a30", bg: "#faece7" };
  const salesCard = { label: "มูลค่าขายสะสม", value: fmt(totalSalesValue), suffix: "บาท", icon: ArrowUpFromLine, color: "#185fa5", bg: "#e6f1fb" };
  const expensesCard = { label: "ค่าใช้จ่ายรวม", value: fmt(totalExpenses), suffix: "บาท", icon: Receipt, color: "#993c1d", bg: "#faece7" };
  const stockCard = {
    label: "ยอดคงเหลือสต็อก (ต้นทุนก่อน VAT)",
    value: fmt(totalStockValue), suffix: "บาท", icon: Boxes, color: "#1d9e75", bg: "#e1f5ee"
  };
  const loanCard = { label: "คงเหลือสินเชื่อ/เงินกู้", value: fmt(totalLoanRemaining), suffix: "บาท", icon: CreditCard, color: "#993c1d", bg: "#faece7" };

  const subTabs = [
    { key: "purchases", label: "ซื้อ", icon: ArrowDownToLine },
    { key: "sales", label: "ขาย", icon: ArrowUpFromLine },
    { key: "expenses", label: "ค่าใช้จ่าย", icon: Receipt },
    { key: "stock", label: "สต็อก", icon: Boxes },
    { key: "loans", label: "สินเชื่อ", icon: CreditCard },
    { key: "cashflow", label: "เงินหมุนร้าน", icon: Landmark },
  ];

  const renderCard = (c, snapshot) => {
    const Icon = c.icon;
    return (
      <div key={c.label} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "16px 18px" }}>
        <div style={{ width: 36, height: 36, borderRadius: 8, background: c.bg, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 12 }}>
          <Icon size={18} color={c.color} />
        </div>
        <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>{c.label} {snapshot && <span style={{ color: "#bcb6e0" }}>(ปัจจุบัน)</span>}</div>
        <div style={{ fontSize: 22, fontWeight: 700 }}>{c.value} <span style={{ fontSize: 13, fontWeight: 400, color: "#9ca3af" }}>{c.suffix}</span></div>
      </div>
    );
  };

  // ---------- Export handlers (per sub-tab) ----------
  const periodLabel = dateRange ? `${dateRange.start} ถึง ${dateRange.end}` : "ทั้งหมด";

  const exportHandlers = {
    purchases: {
      pdf: () => printAsPDF("dash-export-purchases", `ยอดซื้อ (${periodLabel})`),
      excel: () => {
        const rows = [
          [`ยอดซื้อ - ${periodLabel}`, "", ""],
          ["", "", ""],
          ["ประเภทสินค้า", "จำนวน", "มูลค่า (บาท)"],
          ...purchaseByType.map((g) => [g.type, g.qty, g.value]),
          ["", "", ""],
          ["สินค้า", "จำนวน", "มูลค่า (บาท)"],
          ...purchaseByProduct.map((g) => [prodName(g.productId), g.qty, g.value]),
        ];
        exportExcel(rows, `ยอดซื้อ_${periodLabel}.xlsx`, "ยอดซื้อ");
      },
      image: () => printAsPDF("dash-export-purchases", `ยอดซื้อ (${periodLabel})`),
    },
    sales: {
      pdf: () => printAsPDF("dash-export-sales", `ยอดขาย (${periodLabel})`),
      excel: () => {
        const rows = [
          [`ยอดขาย - ${periodLabel}`, "", ""],
          ["", "", ""],
          ["ประเภทสินค้า", "จำนวน", "มูลค่า (บาท)"],
          ...salesByType.map((g) => [g.type, g.qty, g.value]),
          ["", "", ""],
          ["สินค้า", "จำนวน", "มูลค่า (บาท)"],
          ...salesByProduct.map((g) => [prodName(g.productId), g.qty, g.value]),
        ];
        exportExcel(rows, `ยอดขาย_${periodLabel}.xlsx`, "ยอดขาย");
      },
      image: () => printAsPDF("dash-export-sales", `ยอดขาย (${periodLabel})`),
    },
    expenses: {
      pdf: () => printAsPDF("dash-export-expenses", `ค่าใช้จ่าย (${periodLabel})`),
      excel: () => {
        const rows = [
          [`ค่าใช้จ่ายรวม - ${periodLabel}`],
          ["ค่าใช้จ่ายรวม (หมวด ค่าใช้จ่าย, ก่อนภาษี)", totalExpenses],
        ];
        exportExcel(rows, `ค่าใช้จ่าย_${periodLabel}.xlsx`, "ค่าใช้จ่าย");
      },
      image: () => printAsPDF("dash-export-expenses", `ค่าใช้จ่าย (${periodLabel})`),
    },
    stock: {
      pdf: () => printAsPDF("dash-export-stock", "สต๊อกคงเหลือ"),
      excel: () => {
        const rows = [
          ["สต๊อกคงเหลือ", "", "", ""],
          ["ประเภทสินค้า", "สินค้า", "คงเหลือ", "มูลค่า (บาท)", "ราคาเฉลี่ย"],
        ];
        stockByType.forEach((g) => {
          const visibleItems = g.items.filter((s) => s.qty > 0);
          if (visibleItems.length === 0) return;
          rows.push([g.type, "", g.qty, g.value, g.avgCost]);
          visibleItems.forEach((s) => rows.push(["", s.name, s.qty, s.totalCost, s.avgCost]));
        });
        rows.push(["", "", "", ""]);
        rows.push(["ผลรวม", "", stockByType.reduce((s, g) => s + g.qty, 0), stockByType.reduce((s, g) => s + g.value, 0)]);
        exportExcel(rows, "สต๊อกคงเหลือ.xlsx", "สต๊อก");
      },
      image: () => printAsPDF("dash-export-stock", "สต๊อกคงเหลือ"),
    },
    loans: {
      pdf: () => printAsPDF("dash-export-loans", "สินเชื่อ/เงินกู้คงเหลือ"),
      excel: () => {
        const rows = [
          ["สินเชื่อ/เงินกู้คงเหลือ"],
          ["ชื่อสัญญา", "เลขที่บิล", "ประเภท", "เงินต้น", "งวดที่ชำระแล้ว", "งวดคงเหลือ"],
          ...(loans || []).map((l) => {
            const paidCount = (l.paidInstallments || []).length;
            return [l.name, l.billNo || "", l.type, l.principal, paidCount, l.totalInstallments - paidCount];
          }),
          ["", "", "", "", ""],
          ["คงเหลือทั้งหมด (บาท)", totalLoanRemaining],
        ];
        exportExcel(rows, "สินเชื่อคงเหลือ.xlsx", "สินเชื่อ");
      },
      image: () => printAsPDF("dash-export-loans", "สินเชื่อ/เงินกู้คงเหลือ"),
    },
  };

  return (
    <div>
      <h2 style={{ margin: "0 0 4px", fontSize: 22, fontWeight: 700 }}>แดชบอร์ดภาพรวม</h2>
      <p style={{ margin: "0 0 12px", color: "#6b7280", fontSize: 14 }}>สรุปข้อมูลการซื้อขายของเก่ารีไซเคิล</p>





      {/* ตัวเลือกช่วงเวลา */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        <div style={{ display: "flex", borderRadius: 8, overflow: "hidden", border: "1px solid #d1d5db" }}>
          {[
            { key: "all", label: "ทั้งหมด" },
            { key: "day", label: "รายวัน" },
            { key: "range", label: "เลือกช่วงวันที่" },
          ].map((opt) => (
            <button
              key={opt.key}
              onClick={() => setPeriodMode(opt.key)}
              style={{
                padding: "8px 16px", border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600,
                background: periodMode === opt.key ? "#1d9e75" : "#fff",
                color: periodMode === opt.key ? "#04342c" : "#6b7280",
              }}
            >
              {opt.label}
            </button>
          ))}
        </div>

        {periodMode === "day" && (
          <input type="date" style={{ ...inputStyle, width: 170 }} value={periodDate} onChange={(e) => setPeriodDate(e.target.value)} />
        )}

        {periodMode === "range" && (
          <>
            <span style={{ fontSize: 13, color: "#6b7280" }}>จาก</span>
            <input type="date" style={{ ...inputStyle, width: 170 }} value={rangeStart} onChange={(e) => setRangeStart(e.target.value)} />
            <span style={{ fontSize: 13, color: "#6b7280" }}>ถึง</span>
            <input type="date" style={{ ...inputStyle, width: 170 }} value={rangeEnd} onChange={(e) => setRangeEnd(e.target.value)} />
          </>
        )}

        {dateRange && (
          <span style={{ fontSize: 13, color: "#6b7280" }}>
            ช่วงข้อมูล: {dateRange.start} ถึง {dateRange.end}
          </span>
        )}
      </div>

      {/* แท็บหมวดหมู่ */}
      <div style={{ display: "flex", gap: 8, marginBottom: 20, overflowX: "auto", overflowY: "hidden", paddingBottom: 4, WebkitOverflowScrolling: "touch" }}>
        {subTabs.map((t) => {
          const Icon = t.icon;
          const active = dashSubTab === t.key;
          return (
            <button
              key={t.key}
              onClick={() => setDashSubTab(t.key)}
              style={{
                display: "flex", alignItems: "center", gap: 6, flexShrink: 0,
                padding: "8px 16px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600,
                border: active ? "1px solid #1d9e75" : "1px solid #d1d5db",
                background: active ? "#e1f5ee" : "#fff",
                color: active ? "#0f6e56" : "#6b7280",
              }}
            >
              <Icon size={15} />
              {t.label}
            </button>
          );
        })}
      </div>

      {/* ===== ซื้อ ===== */}
      {dashSubTab === "purchases" && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 13, color: "#6b7280" }}>ข้อมูลตามช่วงเวลา: {periodLabel}</span>
            <ExportToolbar onPDF={exportHandlers.purchases.pdf} onExcel={exportHandlers.purchases.excel} onImage={exportHandlers.purchases.image} />
          </div>
          <div id="dash-export-purchases">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14, marginBottom: 20 }}>
            {renderCard(purchaseCard)}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "18px 20px", overflowX: "auto" }}>
  <h3 style={{ margin: "0 0 14px", fontSize: 15, fontWeight: 600 }}>ยอดซื้อ แบ่งตามประเภทสินค้า</h3>
  <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                <colgroup>
                  <col style={{ width: "50%" }} />
                  <col style={{ width: "20%" }} />
                  <col style={{ width: "15%" }} />
                  <col style={{ width: "15%" }} />
                </colgroup>
                <thead>
                  <tr>
                    <th style={thStyle}>ประเภท</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ย/หน่วย</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>มูลค่ารวม</th>
                  </tr>
                </thead>
                <tbody>
                  {purchaseByType.map((g) => (
                    <tr key={g.type}>
                      <td style={tdStyle}><Badge text={g.type} /></td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.qty)}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.avgCost)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(g.value)}</td>
                    </tr>
                  ))}
                  {purchaseByType.length === 0 && <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีข้อมูลในช่วงเวลานี้</td></tr>}
                </tbody>
                {purchaseByType.length > 0 && (
                  <tfoot>
                    <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                      <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(purchaseByType.reduce((s, g) => s + g.qty, 0))}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>—</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>฿{fmt(purchaseByType.reduce((s, g) => s + g.value, 0))}</td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>

            <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "18px 20px", overflowX: "auto" }}>
              <h3 style={{ margin: "0 0 14px", fontSize: 15, fontWeight: 600 }}>ยอดซื้อ แบ่งตามรายการสินค้า</h3>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                <colgroup>
                  <col style={{ width: "50%" }} />
                  <col style={{ width: "20%" }} />
                  <col style={{ width: "15%" }} />
                  <col style={{ width: "15%" }} />
                </colgroup>
                <thead>
                  <tr>
                    <th style={thStyle}>สินค้า</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ย/หน่วย</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>มูลค่ารวม</th>
                  </tr>
                </thead>
                <tbody>
                  {purchaseByProduct.map((g) => (
                    <tr key={g.productId}>
                      <td style={tdStyle}>{prodName(g.productId)}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.qty)} {prodUnit(g.productId)}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.avgCost)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(g.value)}</td>
                    </tr>
                  ))}
                  {purchaseByProduct.length === 0 && <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีข้อมูลในช่วงเวลานี้</td></tr>}
                </tbody>
                {purchaseByProduct.length > 0 && (
                  <tfoot>
                    <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                      <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>—</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>—</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>฿{fmt(purchaseByProduct.reduce((s, g) => s + g.value, 0))}</td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          </div>

          {/* สรุปบิลแยกตามช่องทางชำระ */}
          <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "18px 20px", marginTop: 16 }}>
            <h3 style={{ margin: "0 0 14px", fontSize: 15, fontWeight: 600 }}>สรุปบิลแยกตามช่องทางชำระ</h3>
            {(() => {
              // จัดกลุ่ม purchases ในช่วง
              const pos = filteredPurchases;
              // รวบรวม payments ทั้งหมด พร้อม po info
              const entries = [];
              pos.forEach(po => {
                const custN = (customers.find(c=>c.id===po.customerId)?.name) || "-";
                const sub = (po.items||[]).reduce((s,it)=>{
                  const qty = Number(it.qty)||0;
                  const net = it.deductPct!=null||it.deductKg!=null ? qty-(qty*(Number(it.deductPct)||0)/100)-(Number(it.deductKg)||0) : qty-(Number(it.deduct)||0);
                  return s+net*(Number(it.price)||0);
                },0);
                const total = sub + sub*((Number(po.vatRate)||0)/100);
                (po.payments||[]).forEach(p => {
                  const acc = storeBankAccounts.find(a=>a.id===p.fromStoreBankId);
                  const group = p.fromStoreBankId === "DEPOSIT" ? "หักเงินมัดจำ"
                    : acc ? `${acc.bankName} — ${acc.accountNo}`
                    : (p.method || "เงินสด");
                  entries.push({ group, id: po.id, cust: custN, amount: Number(p.amount)||0 });
                });
                if (!(po.payments||[]).length) {
                  entries.push({ group: "ยังไม่ชำระ", id: po.id, cust: custN, amount: total });
                }
              });
              // จัดกลุ่ม
              const grouped = {};
              entries.forEach(e => {
                if (!grouped[e.group]) grouped[e.group] = [];
                grouped[e.group].push(e);
              });
              if (Object.keys(grouped).length === 0) return <p style={{ color: "#9ca3af", fontSize: 13 }}>ไม่มีข้อมูล</p>;
              return Object.entries(grouped).map(([grp, rows]) => (
                <div key={grp} style={{ marginBottom: 16 }}>
                  <div style={{ fontWeight: 700, fontSize: 13, color: "#185fa5", background: "#e6f1fb", padding: "5px 12px", borderRadius: 6, marginBottom: 6 }}>{grp}</div>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <th style={thStyle}>เลขที่บิล</th>
                        <th style={thStyle}>ชื่อลูกค้า</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r,i) => (
                        <tr key={i}>
                          <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{r.id}</td>
                          <td style={tdStyle}>{r.cust}</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#993c1d" }}>฿{fmt(r.amount)}</td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr style={{ background: "#f9fafb" }}>
                        <td colSpan={2} style={{ ...tdStyle, fontWeight: 700 }}>รวม</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>฿{fmt(rows.reduce((s,r)=>s+r.amount,0))}</td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              ));
            })()}
          </div>
          </div>
        </>
      )}

      {/* ===== ขาย ===== */}
      {dashSubTab === "sales" && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 13, color: "#6b7280" }}>ข้อมูลตามช่วงเวลา: {periodLabel}</span>
            <ExportToolbar onPDF={exportHandlers.sales.pdf} onExcel={exportHandlers.sales.excel} onImage={exportHandlers.sales.image} />
          </div>
          <div id="dash-export-sales">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14, marginBottom: 20 }}>
            {renderCard(salesCard)}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "18px 20px", overflowX: "auto" }}>
              <h3 style={{ margin: "0 0 14px", fontSize: 15, fontWeight: 600 }}>ยอดขาย แบ่งตามประเภทสินค้า</h3>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                <thead>
                  <tr>
                    <th style={thStyle}>ประเภท</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ย/หน่วย</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>มูลค่ารวม</th>
                  </tr>
                </thead>
                <tbody>
                  {salesByType.map((g) => (
                    <tr key={g.type}>
                      <td style={tdStyle}><Badge text={g.type} /></td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.qty)}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.avgCost)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(g.value)}</td>
                    </tr>
                  ))}
                  {salesByType.length === 0 && <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีข้อมูลในช่วงเวลานี้</td></tr>}
                </tbody>
                {salesByType.length > 0 && (
                  <tfoot>
                    <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                      <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(salesByType.reduce((s, g) => s + g.qty, 0))}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>—</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#185fa5" }}>฿{fmt(salesByType.reduce((s, g) => s + g.value, 0))}</td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>

            <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "18px 20px", overflowX: "auto" }}>
              <h3 style={{ margin: "0 0 14px", fontSize: 15, fontWeight: 600 }}>ยอดขาย แบ่งตามรายการสินค้า</h3>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500  }}>
                <thead>
                  <tr>
                    <th style={thStyle}>สินค้า</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ย/หน่วย</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>มูลค่ารวม</th>
                  </tr>
                </thead>
                <tbody>
                  {salesByProduct.map((g) => (
                    <tr key={g.productId}>
                      <td style={tdStyle}>{prodName(g.productId)}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.qty)} {prodUnit(g.productId)}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.avgCost)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(g.value)}</td>
                    </tr>
                  ))}
                  {salesByProduct.length === 0 && <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีข้อมูลในช่วงเวลานี้</td></tr>}
                </tbody>
                {salesByProduct.length > 0 && (
                  <tfoot>
                    <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                      <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>—</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>—</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#185fa5" }}>฿{fmt(salesByProduct.reduce((s, g) => s + g.value, 0))}</td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>
          </div>
          </div>
        </>
      )}

      {/* ===== ค่าใช้จ่าย ===== */}
      {dashSubTab === "expenses" && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 13, color: "#6b7280" }}>ข้อมูลตามช่วงเวลา: {periodLabel}</span>
            <ExportToolbar onPDF={exportHandlers.expenses.pdf} onExcel={exportHandlers.expenses.excel} onImage={exportHandlers.expenses.image} />
          </div>
          <div id="dash-export-expenses">
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14, marginBottom: 20 }}>
              {renderCard(expensesCard)}
            </div>
            <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "18px 20px", overflowX: "auto" }}>
              <h3 style={{ margin: "0 0 14px", fontSize: 15, fontWeight: 600 }}>ค่าใช้จ่าย แบ่งตามหมวดหมู่ย่อย</h3>
              <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                <thead>
                  <tr>
                    <th style={thStyle}>หมวดหมู่ย่อย</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>จำนวนรายการ</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>ยอดรวม (บาท)</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>% ของทั้งหมด</th>
                  </tr>
                </thead>
                <tbody>
                  {expensesBySubCategory.map((g) => (
                    <tr key={g.subCategory}>
                      <td style={tdStyle}>{g.subCategory}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>{g.count} รายการ</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(g.amount)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280" }}>
                        {totalExpenses > 0 ? `${((g.amount / totalExpenses) * 100).toFixed(1)}%` : "-"}
                      </td>
                    </tr>
                  ))}
                  {expensesBySubCategory.length === 0 && <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีข้อมูลในช่วงเวลานี้</td></tr>}
                </tbody>
                {expensesBySubCategory.length > 0 && (
                  <tfoot>
                    <tr>
                      <td style={{ ...tdStyle, fontWeight: 700 }}>รวม</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{expensesBySubCategory.reduce((s, g) => s + g.count, 0)} รายการ</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>฿{fmt(totalExpenses)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>100%</td>
                    </tr>
                  </tfoot>
                )}
              </table>
            </div>

            {/* สรุปค่าใช้จ่ายแยกตามช่องทางชำระ */}
            <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "18px 20px", marginTop: 16 }}>
              <h3 style={{ margin: "0 0 14px", fontSize: 15, fontWeight: 600 }}>สรุปบิลแยกตามช่องทางชำระ</h3>
              {(() => {
                const exps = (expenses || []).filter(e => inRange(e.billDate || e.date));
                const entries = [];
                exps.forEach(e => {
                  const label = e.vendorName || (e.items&&e.items[0]?.subCategory) || "-";
                  const total = (() => {
                    const items = e.items&&e.items.length>0 ? e.items : [{amount:e.amount,vatEnabled:e.vatEnabled,whtRate:e.whtRate}];
                    return items.reduce((s,it)=>{const a=Number(it.amount)||0;return s+a+(it.vatEnabled?a*0.07:0)-a*((Number(it.whtRate)||0)/100);},0);
                  })();
                  (e.payments||[]).forEach(p => {
                    const acc = storeBankAccounts.find(a=>a.id===p.fromStoreBankId);
                    const group = p.fromStoreBankId==="DEPOSIT" ? "หักเงินมัดจำ"
                      : acc ? `${acc.bankName} — ${acc.accountNo}`
                      : (p.method || "เงินสด");
                    entries.push({ group, id: e.refNo||e.id, cust: label, amount: Number(p.amount)||0 });
                  });
                  if (!(e.payments||[]).length) {
                    entries.push({ group: "ยังไม่ชำระ", id: e.refNo||e.id, cust: label, amount: total });
                  }
                });
                const grouped = {};
                entries.forEach(e => { if (!grouped[e.group]) grouped[e.group]=[]; grouped[e.group].push(e); });
                if (Object.keys(grouped).length === 0) return <p style={{ color: "#9ca3af", fontSize: 13 }}>ไม่มีข้อมูล</p>;
                return Object.entries(grouped).map(([grp, rows]) => (
                  <div key={grp} style={{ marginBottom: 16 }}>
                    <div style={{ fontWeight: 700, fontSize: 13, color: "#185fa5", background: "#e6f1fb", padding: "5px 12px", borderRadius: 6, marginBottom: 6 }}>{grp}</div>
                    <table style={{ width: "100%", borderCollapse: "collapse" }}>
                      <colgroup><col style={{ width: "50%" }} /><col style={{ width: "35%" }} /><col style={{ width: "15%" }} /></colgroup>
                      <thead>
                        <tr>
                          <th style={thStyle}>เลขที่บิล</th>
                          <th style={thStyle}>ผู้รับเงิน / รายการ</th>
                          <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((r,i) => (
                          <tr key={i}>
                            <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{r.id}</td>
                            <td style={tdStyle}>{r.cust}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#993c1d" }}>฿{fmt(r.amount)}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: "#f9fafb" }}>
                          <td colSpan={2} style={{ ...tdStyle, fontWeight: 700 }}>รวม</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>฿{fmt(rows.reduce((s,r)=>s+r.amount,0))}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                ));
              })()}
            </div>
          </div>
        </>
      )}

      {/* ===== สต็อก ===== */}
      {dashSubTab === "stock" && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 13, color: "#6b7280" }}>ยอดคงเหลือ ณ วันที่ {today}</span>
            <ExportToolbar onPDF={exportHandlers.stock.pdf} onExcel={exportHandlers.stock.excel} onImage={exportHandlers.stock.image} />
          </div>
          <div id="dash-export-stock">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14, marginBottom: 20 }}>
            {renderCard(stockCard, true)}

          </div>
          <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflowX: "auto" }}>
            <div style={{ background: "#5a1414", color: "#fff", padding: "14px 20px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
              <h3 style={{ margin: 0, fontSize: 17, fontWeight: 700 }}>มูลค่าสต๊อกรวม</h3>
              <span style={{ fontSize: 12, color: "#e7c9c9" }}>วันที่ {today}</span>
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 480 }}>
              <thead>
                <tr>
                  <th style={thStyle}>ประเภทสินค้า / รายการสินค้า</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>คงเหลือ/{stockByType[0]?.items[0]?.unit || "หน่วย"}</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>มูลค่าคงเหลือ</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ย</th>
                </tr>
              </thead>
              <tbody>
                {stockByType.map((g) => {
                  const visibleItems = g.items.filter((s) => s.qty > 0);
                  if (visibleItems.length === 0) return null;
                  const isExpanded = !!expandedStockTypes[g.type];
                  return (
                    <React.Fragment key={g.type}>
                      {isExpanded && (
                        <>
                          <tr
                            onClick={() => setExpandedStockTypes((prev) => ({ ...prev, [g.type]: !prev[g.type] }))}
                            style={{ cursor: "pointer" }}
                            onMouseEnter={(e) => { e.currentTarget.style.background = "#f3f4f6"; }}
                            onMouseLeave={(e) => { e.currentTarget.style.background = ""; }}
                          >
                            <td style={{ ...tdStyle, fontWeight: 700, color: "#993c1d" }}>{g.type}</td>
                            <td style={tdStyle}></td>
                            <td style={tdStyle}></td>
                            <td style={tdStyle}></td>
                          </tr>
                          {visibleItems.map((s) => (
                            <tr key={s.productId}>
                              <td style={{ ...tdStyle, color: "#111827", paddingLeft: 24 }}>- {s.name}</td>
                              <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(s.qty)}</td>
                              <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d" }}>{fmt(s.totalCost)}</td>
                              <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(s.avgCost)}</td>
                            </tr>
                          ))}
                          <tr style={{ background: "#1f2937" }}>
                            <td style={{ ...tdStyle, fontWeight: 700, color: "#fff" }}>{g.type} (ยอดรวม)</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#fff" }}>{fmt(g.qty)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#fca5a5" }}>{fmt(g.value)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#fff" }}>{fmt(g.avgCost)}</td>
                          </tr>
                        </>
                      )}
                      {!isExpanded && (
                        <tr
                          onClick={() => setExpandedStockTypes((prev) => ({ ...prev, [g.type]: !prev[g.type] }))}
                          style={{ cursor: "pointer", background: "#1f2937" }}
                          onMouseEnter={(e) => { e.currentTarget.style.background = "#374151"; }}
                          onMouseLeave={(e) => { e.currentTarget.style.background = "#1f2937"; }}
                        >
                          <td style={{ ...tdStyle, fontWeight: 700, color: "#fff" }}>{g.type} (ยอดรวม)</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#fff" }}>{fmt(g.qty)}</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#fca5a5" }}>{fmt(g.value)}</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#fff" }}>{fmt(g.avgCost)}</td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
                {stockByType.every((g) => g.items.every((s) => s.qty === 0)) && (
                  <tr><td colSpan={4} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีข้อมูลสต๊อก</td></tr>
                )}
              </tbody>
              {stockByType.length > 0 && (
                <tfoot>
                  <tr style={{ borderTop: "3px solid #0f6e56" }}>
                    <td style={{ ...tdStyle, fontWeight: 700, color: "#0f6e56", fontSize: 14 }}>ผลรวม</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56", fontSize: 14 }}>{fmt(stockByType.reduce((s, g) => s + g.qty, 0))}</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56", fontSize: 14 }}>{fmt(stockByType.reduce((s, g) => s + g.value, 0))}</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56", fontSize: 14 }}>
                      {(() => {
                        const totalQty = stockByType.reduce((s, g) => s + g.qty, 0);
                        const totalVal = stockByType.reduce((s, g) => s + g.value, 0);
                        return fmt(totalQty > 0 ? totalVal / totalQty : 0);
                      })()}
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
          </div>{/* end dash-export-stock */}
        </>
      )}

      {/* ===== สินเชื่อ ===== */}
      {dashSubTab === "loans" && (
        <>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 13, color: "#6b7280" }}>ยอดคงเหลือ ณ วันที่ {today}</span>
            <ExportToolbar onPDF={exportHandlers.loans.pdf} onExcel={exportHandlers.loans.excel} onImage={exportHandlers.loans.image} />
          </div>
          <div id="dash-export-loans" style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 }}>
            {renderCard(loanCard, true)}
            {totalOpeningBankBalance > 0 && (
              <div style={{ background: "#e6f1fb", borderRadius: 12, border: "1px solid #b3d0f0", padding: "14px 18px" }}>
                <div style={{ fontSize: 12, color: "#0c447c", marginBottom: 4, fontWeight: 600 }}>ยอดธนาคารยกมา</div>
                <div style={{ fontWeight: 700, fontSize: 18, color: "#185fa5" }}>฿{fmt(totalOpeningBankBalance)}</div>
                <div style={{ fontSize: 11, color: "#6b7280", marginTop: 4 }}>
                  {storeBankAccounts.filter(a => Number(a.openingBalance) > 0).map(a => (
                    <div key={a.id}>{a.bankName}: ฿{fmt(a.openingBalance)}</div>
                  ))}
                </div>
              </div>
            )}
        </div>{/* end dash-export-loans */}
        </>
      )}

      {dashSubTab === "cashflow" && (() => {
        // ===== ตัวกรองช่วงเวลาเฉพาะของ "เงินหมุนร้าน" =====
        const bankInflowsRange = {};
        sales.forEach((inv) => (inv.payments || []).forEach((p) => {
          if (p.toStoreBankId && p.toStoreBankId !== "CASH" && p.toStoreBankId !== "PREPAYMENT" && inRange(p.date)) {
            bankInflowsRange[p.toStoreBankId] = (bankInflowsRange[p.toStoreBankId] || 0) + (Number(p.amount) || 0);
          }
        }));
        (prepayments || []).forEach((p) => {
          if (p.toStoreBankId && p.toStoreBankId !== "CASH" && inRange(p.date)) {
            bankInflowsRange[p.toStoreBankId] = (bankInflowsRange[p.toStoreBankId] || 0) + (Number(p.amount) || 0);
          }
        });
        (bankTransfers || []).forEach((t) => {
          if (t.toBankId && inRange(t.date)) {
            bankInflowsRange[t.toBankId] = (bankInflowsRange[t.toBankId] || 0) + (Number(t.amount) || 0);
          }
        });

        const bankOutflowsRange = {};
        const addOutflowRange = (bankId, amount, date) => {
          if (!bankId || bankId === "CASH" || bankId === "DEPOSIT" || bankId === "PREPAYMENT") return;
          if (!inRange(date)) return;
          bankOutflowsRange[bankId] = (bankOutflowsRange[bankId] || 0) + amount;
        };
        purchases.forEach((po) => (po.payments || []).forEach((p) => addOutflowRange(p.fromStoreBankId, Number(p.amount) || 0, p.date)));
        (deposits || []).forEach((d) => addOutflowRange(d.fromStoreBankId, Number(d.amount) || 0, d.date));
        (expenses || []).forEach((e) => (e.payments || []).forEach((p) => addOutflowRange(p.fromStoreBankId, Number(p.amount) || 0, p.date || e.billDate || e.date)));
        (bankTransfers || []).forEach((t) => addOutflowRange(t.fromBankId, Number(t.amount) || 0, t.date));

        const beforeRangeBalance = {};
        if (dateRange) {
          const beforeDate = (d) => d < dateRange.start;
          storeBankAccounts.forEach((b) => {
            let bal = Number(b.openingBalance) || 0;
            sales.forEach((inv) => (inv.payments || []).forEach((p) => {
              if (p.toStoreBankId === b.id && beforeDate(p.date)) bal += Number(p.amount) || 0;
            }));
            purchases.forEach((po) => (po.payments || []).forEach((p) => {
              if (p.fromStoreBankId === b.id && beforeDate(p.date)) bal -= Number(p.amount) || 0;
            }));
            (deposits || []).forEach((d) => {
              if (d.fromStoreBankId === b.id && beforeDate(d.date)) bal -= Number(d.amount) || 0;
            });
            (expenses || []).forEach((e) => (e.payments || []).forEach((p) => {
              if (p.fromStoreBankId === b.id && beforeDate(p.date || e.billDate || e.date)) bal -= Number(p.amount) || 0;
            }));
            (bankTransfers || []).forEach((t) => {
              if (t.fromBankId === b.id && beforeDate(t.date)) bal -= Number(t.amount) || 0;
              if (t.toBankId === b.id && beforeDate(t.date)) bal += Number(t.amount) || 0;
            });
            beforeRangeBalance[b.id] = bal;
          });
        }

        // 1. เงินในธนาคาร = ยอดยกมา(หรือก่อนช่วง) + รับเข้า(ตามช่วง/ทั้งหมด) - จ่ายออก(ตามช่วง/ทั้งหมด)
        const bankRows = storeBankAccounts.map((b) => {
          const ob      = dateRange ? (beforeRangeBalance[b.id] ?? 0) : (Number(b.openingBalance) || 0);
          const inflow  = dateRange ? (bankInflowsRange[b.id] || 0)  : (bankInflows[b.id] || 0);
          const outflow = dateRange ? (bankOutflowsRange[b.id] || 0) : (bankOutflows[b.id] || 0);
          const balance = ob + inflow - outflow;
          return { ...b, ob, inflow, outflow, balance };
        });
        const totalBankBalance = bankRows.reduce((s, b) => s + b.balance, 0);

        // แบ่งกลุ่มบัญชีตามประเภท (ธนาคาร / เงินสด / ยังไม่ระบุ)
        const bankGroupRows = bankRows.filter((b) => b.accountType === "ธนาคาร");
        const cashGroupRows = bankRows.filter((b) => b.accountType === "เงินสด");
        const unsetGroupRows = bankRows.filter((b) => !b.accountType);
        const bankGroupTotal = bankGroupRows.reduce((s, b) => s + b.balance, 0);
        const cashGroupTotal = cashGroupRows.reduce((s, b) => s + b.balance, 0);
        const unsetGroupTotal = unsetGroupRows.reduce((s, b) => s + b.balance, 0);

        // 2. ลูกหนี้ค้างรับ (ยอดคงค้าง ณ ปัจจุบันเสมอ — ตรงกับหน้ารับ/จ่ายชำระ)
        const totalReceivable = sales.reduce((s, inv) => {
          const subtotal = inv.items.reduce((ss, it) => ss + (it.net || 0) * (it.price || 0), 0);
          const ad = subtotal - (inv.discount || 0);
          const total = ad + ad * ((inv.vatRate || 0) / 100);
          const paid = (inv.payments || []).reduce((ss, p) => ss + (Number(p.amount) || 0), 0);
          const remaining = total - paid;
          if (inv.writeOff || remaining <= 0.01) return s;
          return s + remaining;
        }, 0);

        // 3. เจ้าหนี้ค้างจ่าย (ยอดคงค้าง ณ ปัจจุบัน — ตรงกับหน้ารับ/จ่ายชำระ)
        const totalPayable = purchases.filter(po => (po.status || "") !== "ยกเลิก").reduce((s, po) => {
          const subtotal = po.items.reduce((ss, it) => ss + (it.net || 0) * (it.price || 0), 0);
          const vat = subtotal * ((Number(po.vatRate) || 0) / 100);
          const total = subtotal + vat;
          const paid = (po.payments || []).reduce((ss, p) => ss + (Number(p.amount) || 0), 0);
          const remaining = total - paid;
          if (po.writeOff || remaining <= 0.01) return s;
          return s + remaining;
        }, 0);

        // 4. เงินมัดจำคงเหลือ (ยอดคงค้าง ณ ปัจจุบัน)
        // รวมยอดยกมา (c.depositOpening) + มัดจำที่จ่ายเพิ่ม (deposits array) - มัดจำที่ถูกหักในใบรับสินค้า
        const totalDeposit = (() => {
          const openingTotal = customers.reduce((s, c) => s + (Number(c.depositOpening) || 0), 0);
          const newGiven = (deposits || []).reduce((s, d) => s + (Number(d.amount) || 0), 0);
          const used = purchases.reduce((s, po) => s + (po.payments || [])
            .filter((p) => p.fromStoreBankId === "DEPOSIT")
            .reduce((s2, p) => s2 + (Number(p.amount) || 0), 0), 0);
          return Math.max(0, openingTotal + newGiven - used);
        })();

        // รับล่วงหน้าคงเหลือ (ลูกค้าจ่ายให้ร้านล่วงหน้า)
        const totalPrepayment = (() => {
          const openingTotal = customers.reduce((s, c) => s + (Number(c.prepaymentOpening) || 0), 0);
          const received = (prepayments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
          const used = sales.reduce((s, inv) => s + (inv.payments || [])
            .filter((p) => p.fromStoreBankId === "PREPAYMENT" || p.toStoreBankId === "PREPAYMENT")
            .reduce((s2, p) => s2 + (Number(p.amount) || 0), 0), 0);
          return Math.max(0, openingTotal + received - used);
        })();

        // 5. สต๊อกสินค้า (มูลค่าทุน ณ ปัจจุบัน)
        const stockVal = inventory.summary.reduce((s, x) => s + x.totalCost, 0);

        // เงินหมุนยอดทั้งหมด = ธนาคาร + เงินสด + เงินมัดจำ + ลูกหนี้ - เจ้าหนี้
        const grandTotal = bankGroupTotal + cashGroupTotal + totalDeposit + totalPrepayment + totalReceivable - totalPayable;

        const cfCard = (label, value, color, bg, sub) => (
          <div style={{ background: bg, borderRadius: 12, padding: "14px 18px", border: `1px solid ${color}33` }}>
            <div style={{ fontSize: 12, color, marginBottom: 4, fontWeight: 600 }}>{label}</div>
            <div style={{ fontWeight: 700, fontSize: 20, color }}>฿{fmt(value)}</div>
            {sub && <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 3 }}>{sub}</div>}
          </div>
        );

        return (
          <>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
              <span style={{ fontSize: 13, color: "#6b7280" }}>
                {dateRange ? `ข้อมูลรับเข้า/จ่ายออกตามช่วงเวลา: ${periodLabel} (ยอดยกมา = คงเหลือก่อนวันที่ ${dateRange.start})` : `ยอดเงินหมุนเวียนสะสมทั้งหมด ณ วันที่ ${today}`}
              </span>
              <ExportToolbar
                onPDF={() => printAsPDF("dash-cashflow", "สรุปเงินหมุนร้าน")}
                onExcel={() => {
                  const rows = [
                    ["สรุปเงินหมุนร้าน", dateRange ? periodLabel : "ทั้งหมด"],
                    ["รายการ", "ยอด (บาท)"],
                    ...bankRows.map(b => [`${b.accountType || "ยังไม่ระบุประเภท"} ${b.bankName} ${b.accountNo}`, b.balance]),
                    ["รวมเงินในธนาคาร", bankGroupTotal],
                    ["รวมเงินสด", cashGroupTotal],
                    ["ลูกหนี้ค้างรับ (บวก)", totalReceivable],
                    ["เจ้าหนี้ค้างจ่าย (ลบ)", -totalPayable],
                    ["เงินมัดจำคงเหลือ", totalDeposit],
                    ["รับล่วงหน้าคงเหลือ", totalPrepayment],
                    ["มูลค่าสต๊อก (ทุน)", stockVal],
                    ["เงินหมุนยอดทั้งหมด (ธนาคาร + เงินสด + เงินมัดจำ + ลูกหนี้ - เจ้าหนี้)", grandTotal],
                  ];
                  exportExcel(rows, "เงินหมุนร้าน.xlsx", "เงินหมุน");
                }}
                onImage={() => printAsPDF("dash-cashflow", "สรุปเงินหมุนร้าน")}
              />
            </div>

            <div id="dash-cashflow">
              {/* การ์ดสรุป */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12, marginBottom: 14 }}>
                {cfCard(dateRange ? "เงินในธนาคารรวม (ช่วงที่เลือก)" : "เงินในธนาคารรวม", bankGroupTotal, "#185fa5", "#e6f1fb", `${bankGroupRows.length} บัญชี`)}
                {cfCard("ลูกหนี้ค้างรับ", totalReceivable, "#0f6e56", "#e1f5ee", "รอรับชำระ (ปัจจุบัน)")}
                {cfCard("เจ้าหนี้ค้างจ่าย", totalPayable, "#993c1d", "#faece7", "รอจ่ายชำระ (ปัจจุบัน)")}
                {cfCard("เงินมัดจำคงเหลือ", totalDeposit, "#854f0b", "#faeeda", "มัดจำที่ยังไม่ใช้ (ปัจจุบัน)")}
                {cfCard("รับล่วงหน้าคงเหลือ", totalPrepayment, "#1d4ed8", "#eff6ff", "ลูกค้าจ่ายล่วงหน้าที่ยังไม่ได้ตัด")}
                {cfCard("มูลค่าสต๊อก (ทุน)", stockVal, "#1d9e75", "#e1f5ee", "สินค้าคงเหลือ (ปัจจุบัน)")}
                {cfCard(dateRange ? "เงินสดรวม (ช่วงที่เลือก)" : "เงินสดรวม", cashGroupTotal, "#0f6e56", "#e1f5ee", `${cashGroupRows.length} บัญชี`)}
              </div>

              {/* กรอบสรุปยอดเงินหมุนทั้งหมด — ใหญ่ที่สุด รวมทุกประเภท */}
              <div style={{ background: grandTotal >= 0 ? "#e1f5ee" : "#fcebeb", borderRadius: 16, padding: "24px 28px", border: `3px solid ${grandTotal >= 0 ? "#0f6e56" : "#a32d2d"}`, marginBottom: 20 }}>
                <div style={{ fontSize: 14, color: grandTotal >= 0 ? "#0f6e56" : "#a32d2d", marginBottom: 6, fontWeight: 700 }}>เงินหมุนยอดทั้งหมด</div>
                <div style={{ fontWeight: 700, fontSize: 32, color: grandTotal >= 0 ? "#0f6e56" : "#a32d2d" }}>฿{fmt(grandTotal)}</div>
                <div style={{ fontSize: 12, color: "#6b7280", marginTop: 6 }}>ธนาคาร + เงินสด + เงินมัดจำ + ลูกหนี้ − เจ้าหนี้</div>
              </div>

              {/* ตารางรายละเอียดธนาคาร */}
              <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflowX: "auto", marginBottom: 14 }}>
                <div style={{ background: "#185fa5", color: "#fff", padding: "12px 16px", fontWeight: 700, fontSize: 14 }}>
                  ยอดเงินในธนาคารแต่ละบัญชี{dateRange ? ` — ${periodLabel}` : ""}
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse", tableLayout: "fixed" }}>
                  <thead><tr>
                    <th style={{ ...thStyle, width: "28%" }}>ธนาคาร</th>
                    <th style={{ ...thStyle, width: "18%" }}>เลขบัญชี</th>
                    <th style={{ ...thStyle, textAlign: "right", width: "14%" }}>{dateRange ? "ยอดยกมา (ก่อนช่วง)" : "ยอดยกมา"}</th>
                    <th style={{ ...thStyle, textAlign: "right", color: "#0f6e56", width: "13%" }}>รับเข้า</th>
                    <th style={{ ...thStyle, textAlign: "right", color: "#993c1d", width: "13%" }}>จ่ายออก</th>
                    <th style={{ ...thStyle, textAlign: "right", width: "14%" }}>คงเหลือ</th>
                  </tr></thead>
                  <tbody>
                    {bankGroupRows.length > 0 && (
                      <tr style={{ background: "#fff" }}>
                        <td colSpan={6} style={{ ...tdStyle, fontWeight: 700, color: "#185fa5", display: "flex", alignItems: "center", gap: 6 }}>
                          <Landmark size={13} /> ธนาคาร
                        </td>
                      </tr>
                    )}
                    {bankGroupRows.map((b) => (
                      <tr key={b.id}>
                        <td style={{ ...tdStyle, fontWeight: 600, paddingLeft: 24 }}>{b.bankName}</td>
                        <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{b.accountNo}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280" }}>{b.ob !== 0 ? `฿${fmt(b.ob)}` : "-"}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#0f6e56", fontWeight: 600 }}>฿{fmt(b.inflow)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d", fontWeight: 600 }}>฿{fmt(b.outflow)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 14, color: b.balance >= 0 ? "#185fa5" : "#a32d2d" }}>฿{fmt(b.balance)}</td>
                      </tr>
                    ))}


                    {cashGroupRows.length > 0 && (
                      <tr style={{ background: "#fff" }}>
                        <td colSpan={6} style={{ ...tdStyle, fontWeight: 700, color: "#0f6e56", display: "flex", alignItems: "center", gap: 6 }}>
                          <Wallet size={13} /> เงินสด
                        </td>
                      </tr>
                    )}
                    {cashGroupRows.map((b) => (
                      <tr key={b.id}>
                        <td style={{ ...tdStyle, fontWeight: 600, paddingLeft: 24 }}>{b.bankName}</td>
                        <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{b.accountNo}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280" }}>{b.ob !== 0 ? `฿${fmt(b.ob)}` : "-"}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#0f6e56", fontWeight: 600 }}>฿{fmt(b.inflow)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d", fontWeight: 600 }}>฿{fmt(b.outflow)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 14, color: b.balance >= 0 ? "#185fa5" : "#a32d2d" }}>฿{fmt(b.balance)}</td>
                      </tr>
                    ))}


                    {unsetGroupRows.length > 0 && (
                      <tr style={{ background: "#f3f4f6" }}>
                        <td colSpan={6} style={{ ...tdStyle, fontWeight: 700, color: "#854f0b" }}>
                          ยังไม่ระบุประเภท
                        </td>
                      </tr>
                    )}
                    {unsetGroupRows.map((b) => (
                      <tr key={b.id}>
                        <td style={{ ...tdStyle, fontWeight: 600, paddingLeft: 24 }}>{b.bankName}</td>
                        <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{b.accountNo}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280" }}>{b.ob !== 0 ? `฿${fmt(b.ob)}` : "-"}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#0f6e56", fontWeight: 600 }}>฿{fmt(b.inflow)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d", fontWeight: 600 }}>฿{fmt(b.outflow)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 14, color: b.balance >= 0 ? "#185fa5" : "#a32d2d" }}>฿{fmt(b.balance)}</td>
                      </tr>
                    ))}
                    {unsetGroupRows.length > 0 && (
                      <tr style={{ background: "#f9fafb" }}>
                        <td colSpan={5} style={{ ...tdStyle, fontWeight: 600, color: "#854f0b", paddingLeft: 24 }}>รวมกลุ่มยังไม่ระบุประเภท</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#854f0b" }}>฿{fmt(unsetGroupTotal)}</td>
                      </tr>
                    )}

                    {bankRows.length === 0 && <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีบัญชีธนาคาร</td></tr>}
                  </tbody>
                  {bankRows.length > 0 && (
                    <tfoot>

                      {(() => {
                        const depOpening = customers.reduce((s,c) => s + (Number(c.depositOpening)||0), 0);
                        const depIn = (deposits||[]).reduce((s,d) => s + (Number(d.amount)||0), 0);
                        const depOut = purchases.reduce((s,po) => s + (po.payments||[]).filter(p=>p.fromStoreBankId==="DEPOSIT").reduce((s2,p)=>s2+(Number(p.amount)||0),0), 0);
                        return (
                          <tr style={{ background: "#fff" }}>
                            <td colSpan={2} style={{ ...tdStyle, fontWeight: 700, color: "#854f0b" }}>เงินมัดจำคงเหลือรวม</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#6b7280" }}>฿{fmt(depOpening)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>+฿{fmt(depIn)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>-฿{fmt(depOut)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 15, color: "#854f0b" }}>฿{fmt(totalDeposit)}</td>
                          </tr>
                        );
                      })()}
                      <tr style={{ borderTop: "3px solid #185fa5" }}>
                        <td colSpan={2} style={{ ...tdStyle, fontWeight: 700, color: "#185fa5", fontSize: 14 }}>ยอดรวมทั้งหมด (ธนาคาร + เงินสด + มัดจำ)</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#374151", fontSize: 14 }}>
                          ฿{fmt(bankRows.reduce((s,b)=>s+b.ob,0) + customers.reduce((s,c)=>s+(Number(c.depositOpening)||0),0))}
                        </td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56", fontSize: 14 }}>
                          +฿{fmt(bankRows.reduce((s,b)=>s+b.inflow,0) + (deposits||[]).reduce((s,d)=>s+(Number(d.amount)||0),0))}
                        </td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d", fontSize: 14 }}>
                          -฿{fmt(bankRows.reduce((s,b)=>s+b.outflow,0) + purchases.reduce((s,po)=>s+(po.payments||[]).filter(p=>p.fromStoreBankId==="DEPOSIT").reduce((s2,p)=>s2+(Number(p.amount)||0),0),0))}
                        </td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#185fa5", fontSize: 15 }}>
                          ฿{fmt(totalBankBalance + totalDeposit)}
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>

              {/* ตารางสรุปรวม */}
              <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "hidden" }}>
                <div style={{ background: "#0c443c", color: "#fff", padding: "12px 16px", fontWeight: 700, fontSize: 14 }}>
                  สรุปเงินหมุนเวียนร้าน
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <tbody>
                    {[
                      { label: "เงินในธนาคารรวม", value: bankGroupTotal, color: "#185fa5", sign: "+" },
                      { label: "เงินสดรวม", value: cashGroupTotal, color: "#0f6e56", sign: "+" },
                      { label: "ลูกหนี้การค้า (ค้างรับ)", value: totalReceivable, color: "#0f6e56", sign: "+" },
                      { label: "เจ้าหนี้การค้า (ค้างจ่าย)", value: totalPayable, color: "#993c1d", sign: "−" },
                      { label: "เงินมัดจำคงเหลือ", value: totalDeposit, color: "#854f0b", sign: "+" },
                    ].map((r) => (
                      <tr key={r.label}>
                        <td style={{ ...tdStyle, display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ width: 22, height: 22, borderRadius: "50%", background: r.color + "22", color: r.color, display: "inline-flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 13 }}>{r.sign}</span>
                          {r.label}
                        </td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: r.color }}>฿{fmt(r.value)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr style={{ background: grandTotal >= 0 ? "#e1f5ee" : "#fcebeb", borderTop: "2px solid #0c443c" }}>
                      <td style={{ ...tdStyle, fontWeight: 700, fontSize: 15 }}>เงินหมุนยอดทั้งหมด</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 18, color: grandTotal >= 0 ? "#0f6e56" : "#a32d2d" }}>฿{fmt(grandTotal)}</td>
                    </tr>
                    <tr style={{ background: "#f9fafb" }}>
                      <td style={{ ...tdStyle, color: "#6b7280", fontSize: 12 }}>+ มูลค่าสต๊อกสินค้า (ทุน) — ไม่รวมในเงินสด</td>
                      <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280", fontSize: 12 }}>฿{fmt(stockVal)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>


            </div>
          </>
        );
      })()}
    </div>
  );
}

// ===================================================================
// PRODUCTS TAB
// ===================================================================
function ProductsTab({ products, setProducts, unitOptions, setUnitOptions, productCategories, setProductCategories }) {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [form, setForm] = useState({ id: "", name: "", type: productCategories[0] || "", unit: unitOptions[0] || "กก.", openingQty: 0, openingCost: 0, openingMonth: "", buyPrice: 0, vipPrice: 0 });
  // inline price editing: { productId, field } — ช่องที่กำลังแก้ไขอยู่ในตาราง
  const [inlineEdit, setInlineEdit] = useState(null);
  const [inlineVal, setInlineVal] = useState("");
  const inlineEditRef = useRef(null);
  const inlineValRef = useRef("");

  const startInline = (p, field) => {
    const val = String(Number(p[field]) || 0);
    inlineEditRef.current = { id: p.id, field };
    inlineValRef.current = val;
    setInlineEdit({ id: p.id, field });
    setInlineVal(val);
  };
  const commitInline = async () => {
    const edit = inlineEditRef.current;
    const val = inlineValRef.current;
    if (!edit) return;
    inlineEditRef.current = null;
    setInlineEdit(null);
    const updated = products.map((p) => p.id === edit.id ? { ...p, [edit.field]: Number(val) || 0 } : p);
    setProducts(updated);
    const prod = updated.find((p) => p.id === edit.id);
    if (prod) await updateProduct(prod);
  };

  const filtered = [...products].filter((p) => p.name.includes(search) || p.id.includes(search) || p.type.includes(search)).reverse();

  const monthLabelOfProduct = (ym) => {
    if (!ym) return "";
    const [y, m] = ym.split("-");
    return `${MONTH_NAMES_TH[Number(m)]} ${y}`;
  };

 const openAdd = () => { setForm({ id: genSeqId("P", products), name: "", type: productCategories[0] || "", unit: unitOptions[0] || "กก.", openingQty: 0, openingCost: 0, openingMonth: "", buyPrice: 0, vipPrice: 0 }); setModal({ mode: "add" }); };
  const openEdit = (item) => { setForm({ openingQty: 0, openingCost: 0, openingMonth: "", buyPrice: 0, vipPrice: 0, ...item }); setModal({ mode: "edit", item }); };

  // เมื่อพิมพ์ประเภทสินค้าใหม่ที่ยังไม่มี ให้เพิ่มเข้าฐานข้อมูล productCategories ทันที (ใช้ได้ทุกเครื่องหลังจากนี้)
  const handleTypeChange = (value) => {
    if (value && !productCategories.includes(value)) {
      setProductCategories([...productCategories, value]);
    }
    setForm({ ...form, type: value });
  };

const remove = async (id) => {
  const prev = products;
  setProducts(products.filter((p) => p.id !== id)); // อัปเดตหน้าจอทันที (optimistic)
  const { error } = await deleteProduct(id);
  if (error) {
    alert("ลบสินค้าไม่สำเร็จ กรุณาลองใหม่");
    setProducts(prev); // ย้อนกลับถ้าลบไม่สำเร็จ
  }
};

const save = async () => {
  if (!form.name.trim()) return;
  const cleaned = { ...form, openingQty: Number(form.openingQty) || 0, openingCost: Number(form.openingCost) || 0, buyPrice: Number(form.buyPrice) || 0, vipPrice: Number(form.vipPrice) || 0 };

  if (modal.mode === "add") {
    setProducts([...products, cleaned]); // อัปเดตหน้าจอทันที
    const { error } = await insertProduct(cleaned);
    if (error) {
      alert("บันทึกสินค้าไม่สำเร็จ กรุณาลองใหม่ (อาจมีรหัสสินค้านี้อยู่แล้ว)");
      setProducts(products); // ย้อนกลับ
      return;
    }
  } else {
    setProducts(products.map((p) => (p.id === modal.item.id ? cleaned : p)));
    const { error } = await updateProduct(cleaned);
    if (error) {
      alert("แก้ไขสินค้าไม่สำเร็จ กรุณาลองใหม่");
      return;
    }
  }
  setModal(null);
};

  const totalOpeningValue = products.reduce((s, p) => s + (Number(p.openingQty) || 0) * (Number(p.openingCost) || 0), 0);

 return (
    <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 80px)" }}>
      <div style={{ flexShrink: 0 }}>
      <Header title="ข้อมูลสินค้า (Product Master)" subtitle="ฐานข้อมูลสินค้า — ระบุยอดยกมาเพื่อให้สต๊อกเริ่มต้นถูกต้อง">
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <ExportToolbar
            onPDF={() => printAsPDF("products-print", "ข้อมูลสินค้า")}
            onExcel={() => {
              const rows = [
                ["รหัส", "ชื่อสินค้า", "ประเภท", "หน่วย", "ยอดยกมา (จำนวน)", "ต้นทุน/หน่วย", "มูลค่ายกมา (บาท)"],
                ...products.map((p) => [p.id, p.name, p.type, p.unit,
                  Number(p.openingQty) || 0, Number(p.openingCost) || 0,
                  (Number(p.openingQty) || 0) * (Number(p.openingCost) || 0)]),
                ["", "", "", "", "", "รวมมูลค่ายกมา", totalOpeningValue],
              ];
              exportExcel(rows, "ข้อมูลสินค้า_ยอดยกมา.xlsx", "สินค้า");
            }}
            onImage={() => printAsPDF("products-print", "ข้อมูลสินค้า")}
          />
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> เพิ่มสินค้า</button>
        </div>
      </Header>

     <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาชื่อสินค้า, รหัส หรือประเภท..." />
      </div>
      <div id="products-print" style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "auto", flex: 1 }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={thStyle}>รหัส</th>
              <th style={thStyle}>ชื่อสินค้า</th>
              <th style={thStyle}>ประเภท</th>
              <th style={thStyle}>หน่วย</th>
              <th style={{ ...thStyle, textAlign: "right", color: "#854f0b" }}>ราคาหน้าร้าน/หน่วย</th>
              <th style={{ ...thStyle, textAlign: "right", color: "#534ab7" }}>ราคา VIP/หน่วย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ยอดยกมา (จำนวน)</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ต้นทุน/หน่วย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>มูลค่ายกมา</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((p) => {
              const val = (Number(p.openingQty) || 0) * (Number(p.openingCost) || 0);
              const inlinePriceCell = (field, color) => {
                const isEditing = inlineEdit?.id === p.id && inlineEdit?.field === field;
                const price = Number(p[field]) || 0;
                if (isEditing) return (
                  <input
                    type="number" autoFocus
                    style={{ ...inputStyle, textAlign: "right", width: "100%", fontWeight: 700 }}
                    value={inlineVal}
                    onChange={(e) => { setInlineVal(e.target.value); inlineValRef.current = e.target.value; }}
                    onBlur={commitInline}
                    onKeyDown={(e) => { if (e.key === "Enter") commitInline(); if (e.key === "Escape") setInlineEdit(null); }}
                  />
                );
                return (
                  <div
                    onClick={() => startInline(p, field)}
                    title="คลิกเพื่อแก้ไขราคา"
                    style={{ cursor: "pointer", fontWeight: 700, color: price > 0 ? color : "#d1d5db", padding: "2px 4px", borderRadius: 4, userSelect: "none" }}
                  >
                    {price > 0 ? `฿${fmt(price)}` : <span style={{ fontSize: 12 }}>กดเพื่อกรอก</span>}
                  </div>
                );
              };
              return (
                <tr key={p.id}>
                  <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{p.id}</td>
                  <td style={{ ...tdStyle, fontWeight: 600 }}>{p.name}</td>
                  <td style={tdStyle}><Badge text={p.type} /></td>
                  <td style={tdStyle}>{p.unit}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{inlinePriceCell("buyPrice", "#854f0b")}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{inlinePriceCell("vipPrice", "#534ab7")}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(p.openingQty || 0)}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(p.openingCost || 0)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#185fa5" }}>
                    {val > 0 ? `฿${fmt(val)}` : "-"}
                    {val > 0 && p.openingMonth && (
                      <div style={{ fontSize: 11, color: "#9ca3af", fontWeight: 400 }}>({monthLabelOfProduct(p.openingMonth)})</div>
                    )}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <button style={iconBtn} onClick={() => openEdit(p)}><Edit2 size={14} /></button>
                      <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบสินค้า "${p.name}" ใช่หรือไม่?`, () => remove(p.id))}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && <tr><td colSpan={10} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่พบสินค้า</td></tr>}
          </tbody>
          {products.length > 0 && (
            <tfoot>
              <tr>
                <td colSpan={6} style={{ ...tdStyle, fontWeight: 700 }}>รวมมูลค่ายกมาทั้งหมด</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#185fa5" }}>฿{fmt(totalOpeningValue)}</td>
                <td style={tdStyle}></td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>


      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มสินค้า" : "แก้ไขสินค้า"} onClose={() => setModal(null)}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            <Field label="รหัสสินค้า"><input style={inputStyle} value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} disabled={modal.mode === "edit"} /></Field>
            <Field label="ชื่อสินค้า"><input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
            <Field label="ประเภท">
  <input style={inputStyle} list="product-type-options" value={form.type} onChange={(e) => handleTypeChange(e.target.value)} placeholder="เลือกหรือพิมพ์ประเภทใหม่" />
  <datalist id="product-type-options">
    {productCategories.map((t) => <option key={t} value={t} />)}
  </datalist>
</Field>
            <Field label="หน่วย">
              <input
                style={inputStyle}
                list="unit-options"
                value={form.unit}
                onChange={(e) => setForm({ ...form, unit: e.target.value })}
                onBlur={(e) => {
                  const val = e.target.value.trim();
                  if (val && !unitOptions.includes(val)) {
                    setUnitOptions([...unitOptions, val]);
                  }
                }}
                placeholder="เลือกหรือพิมพ์หน่วยใหม่"
              />
              <datalist id="unit-options">
                {unitOptions.map((u) => <option key={u} value={u} />)}
              </datalist>
            </Field>
          </div>
          <div style={{ background: "#fffbeb", borderRadius: 8, padding: "12px 16px", marginTop: 8, border: "1px solid #fde68a" }}>
            <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8, color: "#854f0b" }}>ราคารับซื้อ (กดที่ช่องราคาในตารางเพื่อแก้เร็วขึ้น)</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
              <Field label={`ราคาหน้าร้าน/หน่วย (บาท/${form.unit || "หน่วย"})`}>
                <input type="number" min={0} style={inputStyle} value={form.buyPrice || 0} onChange={(e) => setForm({ ...form, buyPrice: e.target.value })} placeholder="0" />
              </Field>
              <Field label={`ราคา VIP/หน่วย (บาท/${form.unit || "หน่วย"})`}>
                <input type="number" min={0} style={inputStyle} value={form.vipPrice || 0} onChange={(e) => setForm({ ...form, vipPrice: e.target.value })} placeholder="0" />
              </Field>
            </div>
          </div>

          <div style={{ background: "#f0f9f5", borderRadius: 8, padding: "12px 16px", marginTop: 8 }}>
            <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 10, color: "#0f6e56" }}>ยอดคงเหลือยกมา (ก่อนเริ่มใช้ระบบ)</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
              <Field label={`จำนวนยกมา (${form.unit || "หน่วย"})`}>
                <input type="number" min={0} style={inputStyle} value={form.openingQty} onChange={(e) => setForm({ ...form, openingQty: e.target.value })} placeholder="0" />
              </Field>
              <Field label="ต้นทุน/หน่วย (บาท)">
                <input type="number" min={0} style={inputStyle} value={form.openingCost} onChange={(e) => setForm({ ...form, openingCost: e.target.value })} placeholder="0" />
              </Field>
            </div>
            <Field label="ของเดือน">
              <input type="month" style={inputStyle} value={form.openingMonth || ""} onChange={(e) => setForm({ ...form, openingMonth: e.target.value })} />
            </Field>
            <p style={{ fontSize: 11, color: "#0f6e56", margin: "0 0 4px" }}>
              * ระบุเดือนที่ยอดยกมานี้มีผล — รายงาน/แดชบอร์ดของเดือนนั้นๆ จะนับรวมยอดนี้เข้าไปด้วย
            </p>
            {(Number(form.openingQty) > 0 || Number(form.openingCost) > 0) && (
              <div style={{ fontSize: 13, color: "#0f6e56", fontWeight: 600, marginTop: 6 }}>
                มูลค่ายกมา: ฿{fmt((Number(form.openingQty) || 0) * (Number(form.openingCost) || 0))}
              </div>
            )}
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

function CustomersTab({ customers, setCustomers }) {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [viewIdCard, setViewIdCard] = useState(null); // รูปบัตรประชาชนที่เปิดดูเต็ม
  const blank = { id: "", name: "", taxId: "", address: "", phone: "", line: "", email: "", deliveries: 0, bankAccounts: [], idCardImage: "" };
  const [form, setForm] = useState(blank);

  const filtered = [...customers].filter((c) => c.name.includes(search) || c.id.includes(search) || (c.phone || "").includes(search)).reverse();

  const openAdd = () => { setForm({ ...blank, id: genSeqId("C", customers) }); setModal({ mode: "add" }); };
  const openEdit = (item) => { setForm(JSON.parse(JSON.stringify({ ...blank, ...item }))); setModal({ mode: "edit", item }); };

  const save = () => {
    if (!form.name.trim()) return;
    if (modal.mode === "add") setCustomers([...customers, { ...form, deliveries: Number(form.deliveries) || 0 }]);
    else setCustomers(customers.map((c) => (c.id === modal.item.id ? { ...form, deliveries: Number(form.deliveries) || 0 } : c)));
    setModal(null);
  };

  const remove = (id) => setCustomers(customers.filter((c) => c.id !== id));

  // รับรูปภาพบัตรประชาชนและแปลงเป็น base64
  const handleIdCardImage = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => setForm((f) => ({ ...f, idCardImage: ev.target.result }));
    reader.readAsDataURL(file);
  };

  // --- จัดการบัญชีธนาคารของลูกค้า (หลายบัญชี) ---
  const addBankAccount = () => {
    const newAccount = { id: "CB" + Date.now().toString().slice(-6), bankName: BANK_NAMES[0], accountNo: "", accountName: form.name || "" };
    setForm({ ...form, bankAccounts: [...(form.bankAccounts || []), newAccount] });
  };
  const updateBankAccount = (idx, field, value) => {
    const accounts = [...(form.bankAccounts || [])];
    accounts[idx] = { ...accounts[idx], [field]: value };
    setForm({ ...form, bankAccounts: accounts });
  };
  const removeBankAccount = (idx) => {
    setForm({ ...form, bankAccounts: (form.bankAccounts || []).filter((_, i) => i !== idx) });
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 80px)" }}>
      <div style={{ flexShrink: 0 }}>
      <Header title="ข้อมูลลูกค้า" subtitle="รายชื่อลูกค้าและผู้ส่งของรีไซเคิล">
        <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> เพิ่มลูกค้า</button>
      </Header>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหารหัส, ชื่อ หรือเบอร์โทร..." />
      </div>
      <div style={{ flex: 1, overflow: "auto" }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {filtered.map((c) => (
          <div key={c.id} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "14px 18px", display: "flex", alignItems: "center", gap: 14 }}>
            {/* ID Card thumbnail */}
            <div style={{ flexShrink: 0 }}>
              {c.idCardImage ? (
                <img
                  src={c.idCardImage}
                  alt="บัตรประชาชน"
                  onClick={() => setViewIdCard(c)}
                  style={{ width: 64, height: 42, objectFit: "cover", borderRadius: 6, border: "1px solid #e5e7eb", cursor: "pointer" }}
                  title="คลิกเพื่อดูบัตรประชาชน"
                />
              ) : (
                <div style={{ width: 64, height: 42, borderRadius: 6, border: "1.5px dashed #d1d5db", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 2, color: "#d1d5db" }}>
                  <Users size={16} />
                  <span style={{ fontSize: 8 }}>บัตร ปชช.</span>
                </div>
              )}
            </div>

            {/* Main info */}
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: 4 }}>
                <span style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 12, color: "#9ca3af" }}>{c.id}</span>
                <span style={{ fontWeight: 700, fontSize: 15 }}>{c.name}</span>
                {c.phone && <span style={{ fontSize: 12, color: "#6b7280" }}>📞 {c.phone}</span>}
                {c.line && <span style={{ fontSize: 12, color: "#6b7280" }}>Line: {c.line}</span>}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", fontSize: 12, color: "#9ca3af" }}>
                {c.taxId && <span>เลขบัตร/ภาษี: {c.taxId}</span>}
                {c.address && <span>· {c.address}</span>}
                {(c.bankAccounts || []).length > 0 && <span>· {(c.bankAccounts || []).length} บัญชีธนาคาร</span>}
              </div>
            </div>

            {/* Stats + Actions */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
              <div style={{ textAlign: "right" }}>
                <div style={{ fontSize: 11, color: "#9ca3af" }}>ส่งของ</div>
                <div style={{ fontWeight: 700, fontSize: 16, color: "#3c3489" }}>{c.deliveries} ครั้ง</div>
              </div>
              <button style={iconBtn} onClick={() => openEdit(c)} title="แก้ไข"><Edit2 size={15} /></button>
              <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบลูกค้า "${c.name}" ใช่หรือไม่?`, () => remove(c.id))} title="ลบ"><Trash2 size={15} /></button>
            </div>
          </div>
        ))}
        {filtered.length === 0 && <p style={{ color: "#9ca3af" }}>ไม่พบข้อมูลลูกค้า</p>}
      </div>

      {/* รูปบัตรประชาชนเต็มจอ */}
      {viewIdCard && (
        <Modal title={`บัตรประชาชน · ${viewIdCard.name}`} onClose={() => setViewIdCard(null)}>
          <div style={{ textAlign: "center" }}>
            <img src={viewIdCard.idCardImage} alt="บัตรประชาชน" style={{ maxWidth: "100%", maxHeight: 400, borderRadius: 10, border: "1px solid #e5e7eb" }} />
            <div style={{ marginTop: 10, fontSize: 13, color: "#6b7280" }}>{viewIdCard.name} · {viewIdCard.taxId}</div>
          </div>
        </Modal>
      )}

     </div>{/* end scroll area */}
      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มลูกค้าใหม่" : "แก้ไขข้อมูลลูกค้า"} onClose={() => setModal(null)} wide>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            <Field label="รหัสลูกค้า">
              <input style={inputStyle} value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} disabled={modal.mode === "edit"} />
            </Field>
            <Field label="ชื่อลูกค้า / บริษัท">
              <input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} />
            </Field>
            <Field label="เลขบัตรประชาชน / เลขผู้เสียภาษี">
              <input style={inputStyle} value={form.taxId} onChange={(e) => setForm({ ...form, taxId: e.target.value })} />
            </Field>
            <Field label="เบอร์โทร">
              <input style={inputStyle} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} />
            </Field>
            <Field label="Line ID">
              <input style={inputStyle} value={form.line} onChange={(e) => setForm({ ...form, line: e.target.value })} />
            </Field>
            <Field label="Email">
              <input style={inputStyle} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} />
            </Field>
            <Field label="สถิติจำนวนการส่ง (ครั้ง)">
              <input type="number" style={inputStyle} value={form.deliveries} onChange={(e) => setForm({ ...form, deliveries: e.target.value })} />
            </Field>
          </div>
          <Field label="ที่อยู่">
            <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} />
          </Field>

          {/* รูปภาพบัตรประชาชน */}
          <div style={{ marginTop: 8, marginBottom: 16 }}>
            <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 8 }}>รูปภาพบัตรประชาชน</div>
            <div style={{ display: "flex", alignItems: "flex-start", gap: 14 }}>
              <div>
                {form.idCardImage ? (
                  <img src={form.idCardImage} alt="บัตรประชาชน" style={{ width: 160, height: 100, objectFit: "cover", borderRadius: 8, border: "1px solid #e5e7eb" }} />
                ) : (
                  <div style={{ width: 160, height: 100, borderRadius: 8, border: "1.5px dashed #d1d5db", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 6, color: "#9ca3af", background: "#f9fafb" }}>
                    <FileText size={28} />
                    <span style={{ fontSize: 12 }}>ยังไม่มีรูปภาพ</span>
                  </div>
                )}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                <label style={{ ...btnSecondary, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6, fontSize: 13 }}>
                  <Download size={14} /> อัปโหลดรูปบัตรประชาชน
                  <input type="file" accept="image/*" onChange={handleIdCardImage} style={{ display: "none" }} />
                </label>
                {form.idCardImage && (
                  <button style={btnDanger} onClick={() => setForm((f) => ({ ...f, idCardImage: "" }))}>
                    <X size={14} /> ลบรูปภาพ
                  </button>
                )}
                <p style={{ fontSize: 11, color: "#9ca3af", margin: 0 }}>รองรับ JPG, PNG, WEBP (รูปจะเก็บในระบบ)</p>
              </div>
            </div>
          </div>

          <div style={{ marginTop: 8, marginBottom: 8, fontWeight: 600, fontSize: 14 }}>บัญชีธนาคารของลูกค้า (เพิ่มได้หลายบัญชี)</div>
          {(form.bankAccounts || []).length === 0 && <p style={{ color: "#9ca3af", fontSize: 13, marginTop: 0 }}>ยังไม่มีบัญชีธนาคาร</p>}
          {(form.bankAccounts || []).map((b, idx) => (
            <div key={b.id} style={{ display: "grid", gridTemplateColumns: "1.2fr 1.4fr 1.4fr auto", gap: 8, marginBottom: 8, alignItems: "center" }}>
              <select style={inputStyle} value={b.bankName} onChange={(e) => updateBankAccount(idx, "bankName", e.target.value)}>
                {BANK_NAMES.map((n) => <option key={n} value={n}>{n}</option>)}
              </select>
              <input style={inputStyle} placeholder="เลขที่บัญชี" value={b.accountNo} onChange={(e) => updateBankAccount(idx, "accountNo", e.target.value)} />
              <input style={inputStyle} placeholder="ชื่อบัญชี" value={b.accountName} onChange={(e) => updateBankAccount(idx, "accountName", e.target.value)} />
              <button style={btnDanger} onClick={() => removeBankAccount(idx)}><Trash2 size={14} /></button>
            </div>
          ))}
          <button style={btnSecondary} onClick={addBankAccount}><Plus size={14} /> เพิ่มบัญชีธนาคาร</button>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ===================================================================
// PURCHASES TAB (ใบรับสินค้า)
// ===================================================================
function PurchasesTab({ products, customers, purchases, setPurchases, storeBankAccounts, deposits, companySettings }) {
  const [modal, setModal] = useState(null); // {mode:'add'|'edit'|'view', item}
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [expanded, setExpanded] = useState(null);

  const blankItem = () => ({ productId: "", qty: 0, deductPct: 0, deductKg: 0, price: 0 });
  const blankPayment = () => ({
    id: "PM" + Date.now().toString().slice(-6),
    date: new Date().toISOString().slice(0, 10),
    amount: 0,
    fromStoreBankId: storeBankAccounts[0]?.id || "",
    method: PAYMENT_METHODS[0],
  });
  const blankForm = () => ({ id: "", date: new Date().toISOString().slice(0, 10), customerId: "", status: "รออนุมัติ", paymentMethod: PURCHASE_PAYMENT_CHANNELS[0], receivingCustomerBankId: "", items: [blankItem()], payments: [], vatRate: 0, vehiclePlate: "", priceType: "normal" });
  const [form, setForm] = useState(blankForm());

  const custName = (id) => customers.find((c) => c.id === id)?.name || id;
  const prodName = (id) => products.find((p) => p.id === id)?.name || id;
  const prodUnit = (id) => products.find((p) => p.id === id)?.unit || "";
  const custBankAccounts = (customerId) => customers.find((c) => c.id === customerId)?.bankAccounts || [];

  const filtered = purchases.filter((po) => po.id.includes(search) || custName(po.customerId).includes(search)).filter((po) => (!dateFrom || (po.date || "") >= dateFrom) && (!dateTo || (po.date || "") <= dateTo)).sort((a, b) => (b.date || "").localeCompare(a.date || "") || b.id.localeCompare(a.id));

  // ยอดมัดจำคงเหลือของลูกค้าที่เลือก (ไม่รวมยอดที่กำลังหักในใบนี้ จากใบอื่นๆทั้งหมด)
  const depositBalanceForCustomer = (customerId, excludePoId) => {
    const opening = Number(customers.find((c) => c.id === customerId)?.depositOpening) || 0;
    const totalGiven = opening + (deposits || []).filter((d) => d.customerId === customerId).reduce((s, d) => s + (Number(d.amount) || 0), 0);
    const totalUsedOtherPOs = purchases
      .filter((po) => po.customerId === customerId && po.id !== excludePoId)
      .reduce((s, po) => s + (po.payments || []).filter((p) => p.fromStoreBankId === "DEPOSIT").reduce((s2, p) => s2 + (Number(p.amount) || 0), 0), 0);
    return totalGiven - totalUsedOtherPOs;
  };

  const openAdd = () => {
    const _d1 = new Date().toISOString().slice(0, 10);
    setForm({ ...blankForm(), id: genId("PO", purchases, _d1) });
    setModal({ mode: "add" });
  };
  const openEdit = (item) => { setForm(JSON.parse(JSON.stringify({ payments: [], status: "รออนุมัติ", paymentMethod: PURCHASE_PAYMENT_CHANNELS[0], receivingCustomerBankId: "", vatRate: 0, ...item }))); setModal({ mode: "edit", item }); };
  const openView = (item) => setModal({ mode: "view", item });

  const updateItem = (idx, field, value) => {
    const items = [...form.items];
    // เมื่อเลือกสินค้าใหม่ ให้ดึงราคาตามประเภทที่เลือกในใบรับ (หน้าร้าน/VIP)
    if (field === "productId") {
      const prod = products.find((p) => p.id === value);
      let price = items[idx].price;
      if (prod) {
        price = form.priceType === "vip"
          ? (Number(prod.vipPrice) || Number(prod.buyPrice) || price)
          : (Number(prod.buyPrice) || price);
      }
      items[idx] = { ...items[idx], productId: value, price };
    } else {
      items[idx] = { ...items[idx], [field]: value };
    }
    setForm({ ...form, items });
  };
  const addItem = () => setForm({ ...form, items: [...form.items, blankItem()] });
  const removeItem = (idx) => setForm({ ...form, items: form.items.filter((_, i) => i !== idx) });

  const addPayment = () => {
    setForm({ ...form, payments: [...(form.payments || []), blankPayment()] });
  };
  const updatePayment = (idx, field, value) => {
    const payments = [...(form.payments || [])];
    payments[idx] = { ...payments[idx], [field]: value };
    setForm({ ...form, payments });
  };
  const removePayment = (idx) => setForm({ ...form, payments: (form.payments || []).filter((_, i) => i !== idx) });

  const save = () => {
    if (!form.id.trim() || form.items.length === 0) return;
    const cleaned = {
      ...form,
      updated_at: new Date().toISOString(),
      items: form.items.map((it) => {
        const qty = Number(it.qty) || 0;
        const net = lineNet(it);
        const price = Number(it.price) || 0;
        const discountPct = Number(it.discountPct) || 0;
        const discountedPrice = price * (1 - discountPct / 100);
        return { ...it, qty, net, price, discountPct, discountedPrice };
      }),
      payments: (form.payments || []).map((p) => ({ ...p, amount: Number(p.amount) || 0 })),
    };
    if (modal.mode === "add") setPurchases([...purchases, cleaned]);
    else setPurchases(purchases.map((p) => (p.id === modal.item.id ? cleaned : p)));
    setModal(null);
  };

  const remove = (id) => setPurchases(purchases.filter((p) => p.id !== id));

  const approve = (id) => setPurchases(purchases.map((p) => (p.id === id ? { ...p, status: "อนุมัติแล้ว", updated_at: new Date().toISOString() } : p)));
  const cancelPO = (id) => setPurchases(purchases.map((p) => (p.id === id ? { ...p, status: "ยกเลิก" } : p)));
  const revertToPending = (id) => setPurchases(purchases.map((p) => (p.id === id ? { ...p, status: "รออนุมัติ" } : p)));

  const lineNet = (it) => {
    const qty = Number(it.qty) || 0;
    const deductPct = Number(it.deductPct) || 0;
    const deductKg = Number(it.deductKg) || 0;
    // รองรับทั้ง field ใหม่ (deductPct/deductKg) และ field เก่า (deduct/deductType)
    if (it.deductPct != null || it.deductKg != null) {
      return qty - (qty * deductPct / 100) - deductKg;
    }
    const deduct = Number(it.deduct) || 0;
    return it.deductType === "pct" ? qty * (1 - deduct / 100) : qty - deduct;
  };
  const lineTotal = (it) => lineNet(it) * (Number(it.price) || 0) * (1 - (Number(it.discountPct) || 0) / 100);
  const subtotalBeforeVat = (po) => po.items.reduce((s, it) => s + lineNet(it) * (Number(it.price) || 0) * (1 - (Number(it.discountPct) || 0) / 100), 0);
  const vatAmount = (po) => subtotalBeforeVat(po) * ((Number(po.vatRate) || 0) / 100);
  const grandTotal = (po) => subtotalBeforeVat(po) + vatAmount(po);
  const paidTotal = (po) => (po.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);

  const statusBadge = (status) => {
    if (status === "อนุมัติแล้ว") return { bg: "#eaf3de", color: "#27500a", icon: CheckCircle2 };
    if (status === "ยกเลิก") return { bg: "#fcebeb", color: "#791f1f", icon: XCircle };
    return { bg: "#faeeda", color: "#854f0b", icon: Clock };
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 80px)" }}>
      <div style={{ flexShrink: 0 }}>
      <Header title="ใบรับสินค้า (รับซื้อของเก่า)" subtitle="บันทึกการรับซื้อสินค้าจากลูกค้า">
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <ExportToolbar
            onPDF={() => printAsPDF("tab-export-purchases", "ใบรับสินค้า")}
            onExcel={() => {
              const rows = [
                ["เลขที่ใบรับ", "วันที่", "ลูกค้า", "สถานะ", "รายการสินค้า", "จำนวน", "ราคา/หน่วย", "รวม"],
                ...filtered.flatMap((po) =>
                  po.items.map((it, i) => [
                    i === 0 ? po.id : "", i === 0 ? po.date : "", i === 0 ? custName(po.customerId) : "", i === 0 ? po.status : "",
                    prodName(it.productId), it.net, it.price, it.net * it.price,
                  ])
                ),
              ];
              exportExcel(rows, "ใบรับสินค้า.xlsx", "ใบรับสินค้า");
            }}
            onImage={() => printAsPDF("tab-export-purchases", "ใบรับสินค้า")}
          />
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> สร้างใบรับสินค้า</button>
        </div>
      </Header>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาเลขที่ใบรับ หรือชื่อลูกค้า..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />
      </div>
     <div id="tab-export-purchases" style={{ flex: 1, overflow: "auto" }}>
<div style={{ display: "flex", flexDirection: "column", gap: 10, minWidth: 700 }}>
        {filtered.map((po) => {
          const sb = statusBadge(po.status || "รออนุมัติ");
          const SIcon = sb.icon;
          const paid = paidTotal(po);
          const total = grandTotal(po);
          const remaining = total - paid;
          const payBadge = (po.writeOff || remaining <= 0.01) ? { bg: "#eaf3de", color: "#27500a", icon: CheckCircle2, label: "ชำระแล้ว" }
            : paid > 0.01 ? { bg: "#faeeda", color: "#854f0b", icon: Clock, label: "ชำระบางส่วน" }
            : { bg: "#fcebeb", color: "#791f1f", icon: Clock, label: "ค้างจ่าย" };
          const PIcon = payBadge.icon;
          const isExpanded = expanded === po.id;

          return (
            <div key={po.id} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "14px 18px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 600, fontSize: 14, display: "inline-flex", alignItems: "center", gap: 6, color: "#6b7280" }}>
                      <FileText size={14} /> {po.id}
                    </span>
                    <span style={{ fontWeight: 600, fontSize: 14 }}>{custName(po.customerId)}</span>
                    <span style={{ background: sb.bg, color: sb.color, padding: "2px 10px", borderRadius: 6, fontSize: 12, fontWeight: 500, display: "inline-flex", alignItems: "center", gap: 4 }}>
                      <SIcon size={13} /> {po.status || "รออนุมัติ"}
                    </span>
                    <span style={{ background: payBadge.bg, color: payBadge.color, padding: "2px 10px", borderRadius: 6, fontSize: 12, fontWeight: 500, display: "inline-flex", alignItems: "center", gap: 4 }}>
                      <PIcon size={13} /> {payBadge.label}
                    </span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 14, marginTop: 6, fontSize: 13, color: "#6b7280", flexWrap: "wrap" }}>
                    <span>วันที่: {po.date}</span>
                    <span>{po.items.length} รายการสินค้า</span>
                    <span>{po.paymentMethod || "เงินสด"}</span>
                    {po.vehiclePlate && <span style={{ background: "#f3f4f6", padding: "1px 8px", borderRadius: 4, fontSize: 12 }}>🚛 {po.vehiclePlate}</span>}
                  </div>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
                  <div style={{ textAlign: "right" }}>
                    {(po.vatRate > 0) && <div style={{ fontSize: 11, color: "#9ca3af" }}>ก่อน VAT: ฿{fmt(subtotalBeforeVat(po))}</div>}
                    {(po.vatRate > 0) && <div style={{ fontSize: 11, color: "#9ca3af" }}>VAT {po.vatRate}%: +฿{fmt(vatAmount(po))}</div>}
                    <div style={{ fontSize: 12, color: "#9ca3af" }}>ยอดรวม</div>
                    <div style={{ fontSize: 18, fontWeight: 700, color: "#993c1d" }}>฿{fmt(total)}</div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    {(po.status || "รออนุมัติ") === "อนุมัติแล้ว" ? (
                      <button
                        style={{ ...iconBtn, background: "#639922", borderColor: "#639922", color: "#fff" }}
                        onClick={() => revertToPending(po.id)}
                        aria-label="อนุมัติแล้ว (กดเพื่อยกเลิกอนุมัติ)"
                        title="อนุมัติแล้ว — กดเพื่อยกเลิกอนุมัติ"
                      >
                        <CheckCircle2 size={16} />
                      </button>
                    ) : (
                      <button
                        style={{ ...iconBtn, background: "#e24b4a", borderColor: "#e24b4a", color: "#fff" }}
                        onClick={() => approve(po.id)}
                        aria-label="ยังไม่อนุมัติ (กดเพื่ออนุมัติ)"
                        title="ยังไม่อนุมัติ — กดเพื่ออนุมัติ"
                      >
                        <XCircle size={16} />
                      </button>
                    )}
                    <button style={iconBtn} onClick={() => setExpanded(isExpanded ? null : po.id)} aria-label="รายละเอียด" title="ดูรายละเอียด">
                      {isExpanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    </button>
                    <button style={iconBtn} onClick={() => openView(po)} aria-label="พิมพ์ PDF"><Printer size={16} /></button>
                    <button style={iconBtn} onClick={() => openEdit(po)} aria-label="แก้ไข"><Edit2 size={16} /></button>
                    <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบใบรับสินค้า "${po.id}" ใช่หรือไม่?`, () => remove(po.id))} aria-label="ลบ"><Trash2 size={16} /></button>
                  </div>
                </div>
              </div>

              {isExpanded && (
                <div style={{ marginTop: 14, paddingTop: 14, borderTop: "1px solid #f3f4f6" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, marginBottom: 12 }}>
                    <thead>
                      <tr>
                        <th style={thStyle}>สินค้า</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>หัก</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>สุทธิ</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>ราคา/หน่วย</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>หัก %</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
                      </tr>
                    </thead>
                    <tbody>
                      {po.items.map((it, idx) => {
                        const net = calcNet(it);
                        const deductDisplay = (Number(it.qty) || 0) - net;
                        const discountPct = Number(it.discountPct) || 0;
                        const discountedPrice = it.price * (1 - discountPct / 100);
                        return (
                          <tr key={idx}>
                            <td style={tdStyle}>{prodName(it.productId)}</td>
                            <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(it.qty)} {prodUnit(it.productId)}</td>
                            <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(deductDisplay)}</td>
                            <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(net)}</td>
                            <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(it.price)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", color: discountPct > 0 ? "#993c1d" : "#9ca3af" }}>{discountPct > 0 ? `${discountPct}%` : "—"}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>{fmt(net * discountedPrice)}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>

                  <div style={{ background: "#f9fafb", borderRadius: 8, padding: "10px 16px", marginBottom: 12, fontSize: 13, maxWidth: 360 }}>
                    <Row label="ยอดก่อน VAT" value={`฿${fmt(subtotalBeforeVat(po))}`} />
                    {(po.vatRate > 0) && <Row label={`VAT ${po.vatRate}%`} value={`+฿${fmt(vatAmount(po))}`} color="#993c1d" />}
                    <Row label="ยอดรวมที่ต้องชำระ" value={`฿${fmt(total)}`} bold />
                    <Row label="ชำระแล้ว" value={`฿${fmt(paid)}`} />
                    <Row label="คงค้าง" value={`฿${fmt(remaining)}`} bold color={remaining > 0 ? "#a32d2d" : "#27500a"} />
                  </div>

                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {(po.status || "รออนุมัติ") === "รออนุมัติ" && (
                      <button style={{ ...iconBtn, color: "#0f6e56", borderColor: "#9fe1cb" }} onClick={() => approve(po.id)}><CheckCircle2 size={14} /> อนุมัติ</button>
                    )}
                    {(po.status || "รออนุมัติ") === "อนุมัติแล้ว" && (
                      <button style={iconBtn} onClick={() => revertToPending(po.id)}><Clock size={14} /> ยกเลิกอนุมัติ</button>
                    )}
                    {(po.status || "รออนุมัติ") !== "ยกเลิก" && (
                      <button style={btnDanger} onClick={() => cancelPO(po.id)}><XCircle size={14} /> ยกเลิกใบรับ</button>
                    )}
                  </div>
                </div>
              )}
            </div>
          );
        })}
        {filtered.length === 0 && (
          <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "24px", textAlign: "center", color: "#9ca3af" }}>
            ไม่พบใบรับสินค้า
          </div>
        )}
      </div>

      {modal && (modal.mode === "add" || modal.mode === "edit") && (
        <Modal title={`${modal.mode === "add" ? "สร้างใบรับสินค้า" : "แก้ไขใบรับสินค้า"}${modal.mode === "edit" ? " · " + form.id : ""}`} onClose={() => setModal(null)} wide fullscreen>
          <div data-kbform>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1.4fr", gap: "0 16px" }}>
            <Field label="เลขที่ใบรับสินค้า">
              <input style={inputStyle} value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)} />
            </Field>
            <Field label="วันที่ซื้อ">
              <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value, id: genId("PO", purchases, e.target.value) })} onKeyDown={(e) => handleEnterNavigate(e, save)} />
            </Field>
            <Field label="ลูกค้า (ผู้ขาย)">
              <CustomerSelect customers={customers} value={form.customerId} onChange={(cid) => setForm({ ...form, customerId: cid })} />
            </Field>
            <Field label="ประเภทราคา">
              <div style={{ display: "flex", gap: 8, height: 38, alignItems: "center" }}>
                {[{ value: "normal", label: "ราคาหน้าร้าน", color: "#854f0b", bg: "#fffbeb" }, { value: "vip", label: "ราคา VIP", color: "#534ab7", bg: "#f0effe" }].map((opt) => (
                  <button key={opt.value} type="button"
                    style={{ padding: "6px 16px", borderRadius: 6, border: `2px solid ${form.priceType === opt.value ? opt.color : "#e5e7eb"}`, background: form.priceType === opt.value ? opt.bg : "#fff", color: form.priceType === opt.value ? opt.color : "#6b7280", fontWeight: form.priceType === opt.value ? 700 : 400, fontSize: 13, cursor: "pointer" }}
                    onClick={() => {
                      const newItems = form.items.map((it) => {
                        const prod = products.find((p) => p.id === it.productId);
                        if (!prod) return it;
                        const price = opt.value === "vip" ? (Number(prod.vipPrice) || Number(prod.buyPrice) || it.price) : (Number(prod.buyPrice) || it.price);
                        return { ...it, price };
                      });
                      setForm({ ...form, priceType: opt.value, items: newItems });
                    }}
                  >{opt.label}</button>
                ))}
              </div>
            </Field>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1.6fr", gap: "0 16px" }}>
            <Field label="ช่องทางชำระเงิน">
              <select style={inputStyle} value={form.paymentMethod || PURCHASE_PAYMENT_CHANNELS[0]} onChange={(e) => setForm({ ...form, paymentMethod: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)}>
                {PURCHASE_PAYMENT_CHANNELS.map((m) => <option key={m} value={m}>{m}</option>)}
              </select>
            </Field>
            <Field label="บัญชีลูกค้าที่จะรับเงิน">
              <select style={inputStyle} value={form.receivingCustomerBankId || ""} onChange={(e) => setForm({ ...form, receivingCustomerBankId: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)}>
                <option value="">-- เลือกบัญชีลูกค้า --</option>
                {custBankAccounts(form.customerId).map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo} ({b.accountName})</option>)}
              </select>
            </Field>
          </div>

          <div style={{ marginBottom: 6, fontSize: 13, fontWeight: 500, color: "#374151" }}>สถานะใบรับสินค้า</div>
          <div style={{ display: "flex", gap: 10, marginBottom: 6 }}>
            <button
              onClick={() => setForm({ ...form, status: "รออนุมัติ" })}
              style={{
                display: "inline-flex", alignItems: "center", gap: 6, padding: "8px 16px", borderRadius: 8, fontSize: 14, cursor: "pointer",
                border: (form.status || "รออนุมัติ") === "รออนุมัติ" ? "1px solid #f0997b" : "1px solid #d1d5db",
                background: (form.status || "รออนุมัติ") === "รออนุมัติ" ? "#faece7" : "#fff",
                color: (form.status || "รออนุมัติ") === "รออนุมัติ" ? "#993c1d" : "#374151",
                fontWeight: (form.status || "รออนุมัติ") === "รออนุมัติ" ? 600 : 400,
              }}
            >
              <Clock size={15} /> รออนุมัติ
            </button>
            <button
              onClick={() => setForm({ ...form, status: "อนุมัติแล้ว" })}
              style={{
                display: "inline-flex", alignItems: "center", gap: 6, padding: "8px 16px", borderRadius: 8, fontSize: 14, cursor: "pointer",
                border: form.status === "อนุมัติแล้ว" ? "1px solid #5dcaa5" : "1px solid #d1d5db",
                background: form.status === "อนุมัติแล้ว" ? "#e1f5ee" : "#fff",
                color: form.status === "อนุมัติแล้ว" ? "#085041" : "#374151",
                fontWeight: form.status === "อนุมัติแล้ว" ? 600 : 400,
              }}
            >
              <CheckCircle2 size={15} /> อนุมัติแล้ว
            </button>
          </div>
          <p style={{ fontSize: 12, color: "#9ca3af", marginTop: 0, marginBottom: 16 }}>
            ใบที่ "รออนุมัติ" จะยังไม่นำเข้าสต๊อกและไม่ตัดบัญชี จนกว่าจะอนุมัติ
          </p>

          <div style={{ marginTop: 8, marginBottom: 8, fontWeight: 600, fontSize: 14 }}>สินค้า</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 750, tableLayout: "fixed" }}>
              <thead>
                <tr>
                  <th style={thStyle}>สินค้า</th>
                  <th style={{ ...thStyle, textAlign: "right", width: 90 }}>จำนวน</th>
                  <th style={{ ...thStyle, width: 52 }}>หน่วย</th>
                  <th style={{ ...thStyle, textAlign: "right", width: 75 }}>หัก%</th>
                  <th style={{ ...thStyle, textAlign: "right", width: 80 }}>หัก(กก.)</th>
                  <th style={{ ...thStyle, textAlign: "right", width: 80 }}>รวมหัก</th>
                  <th style={{ ...thStyle, textAlign: "right", width: 75 }}>สุทธิ</th>
                  <th style={{ ...thStyle, textAlign: "right", width: 100 }}>ราคา/หน่วย</th>
                  <th style={{ ...thStyle, textAlign: "right", width: 100 }}>จำนวนเงิน</th>
                  <th style={{ ...thStyle, width: 40 }}></th>
                </tr>
              </thead>
              <tbody>
                {form.items.map((it, idx) => {
                  const qty = Math.round((Number(it.qty) || 0) * 100) / 100;
                  const deductPct = Number(it.deductPct) || 0;
                  const deductKg = Math.round((Number(it.deductKg) || 0) * 100) / 100;
                  const totalDeductKg = Math.round(((qty * deductPct / 100) + deductKg) * 100) / 100;
                  const net = Math.round((qty - totalDeductKg) * 100) / 100;
                  return (
                    <tr key={idx}>
                      <td style={tdStyle}>
                        <ProductSelect products={products} value={it.productId} onChange={(pid) => updateItem(idx, "productId", pid)} />
                      </td>
                      <td style={tdStyle}><NumInput style={{ ...inputStyle, width: "100%", textAlign: "right" }} value={it.qty != null ? (Math.round((Number(it.qty)||0)*100)/100) : ""} onChange={(e) => updateItem(idx, "qty", e.target.value)} onKeyDown={(e) => handleEnterNavigate(e, save)} /></td>
                      <td style={{ ...tdStyle, color: "#9ca3af", fontSize: 11 }}>{prodUnit(it.productId)}</td>
                      <td style={tdStyle}><input type="number" min={0} max={100} style={{ ...inputStyle, width: "100%", textAlign: "right" }} value={it.deductPct || ""} placeholder="0" onChange={(e) => updateItem(idx, "deductPct", e.target.value)} /></td>
                      <td style={tdStyle}><input type="number" min={0} style={{ ...inputStyle, width: "100%", textAlign: "right" }} value={it.deductKg != null ? (Math.round((Number(it.deductKg)||0)*100)/100) : ""} placeholder="0" onChange={(e) => updateItem(idx, "deductKg", e.target.value)} /></td>
                      <td style={{ ...tdStyle, textAlign: "right", color: totalDeductKg > 0 ? "#993c1d" : "#9ca3af", fontWeight: 500 }}>{fmt(totalDeductKg)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>{fmt(net)}</td>
                      <td style={tdStyle}><NumInput style={{ ...inputStyle, width: "100%", textAlign: "right" }} value={it.price} onChange={(e) => updateItem(idx, "price", e.target.value)} onKeyDown={(e) => handleEnterNavigate(e, save)} /></td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#993c1d" }}>{fmt(net * (Number(it.price) || 0))}</td>
                      <td style={tdStyle}><button style={btnDanger} onClick={() => removeItem(idx)}><Trash2 size={14} /></button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10, flexWrap: "wrap", gap: 10 }}>
            <button style={btnSecondary} onClick={addItem}><Plus size={14} /> เพิ่มรายการสินค้า</button>
            <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                <label style={{ color: "#6b7280" }}>VAT (%):</label>
                <input
                  type="number"
                  min={0}
                  max={100}
                  style={{ ...inputStyle, width: 80 }}
                  value={form.vatRate}
                  onChange={(e) => setForm({ ...form, vatRate: e.target.value })}
                  onKeyDown={(e) => handleEnterNavigate(e, save)}
                  placeholder="0"
                />
              </div>
              <div style={{ fontSize: 14, fontWeight: 700 }}>
                ก่อน VAT: ฿{fmt(form.items.reduce((s, it) => s + lineTotal(it), 0))}
                {Number(form.vatRate) > 0 && (
                  <span style={{ color: "#993c1d", marginLeft: 10 }}>
                    VAT {form.vatRate}%: +฿{fmt(form.items.reduce((s, it) => s + lineTotal(it), 0) * ((Number(form.vatRate) || 0) / 100))}
                    &nbsp;|&nbsp; รวม: ฿{fmt(form.items.reduce((s, it) => s + lineTotal(it), 0) * (1 + (Number(form.vatRate) || 0) / 100))}
                  </span>
                )}
              </div>
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px", marginTop: 12 }}>
            <Field label="ทะเบียนรถ (ถ้ามี)">
              <input style={inputStyle} value={form.vehiclePlate || ""} onChange={(e) => setForm({ ...form, vehiclePlate: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)} placeholder="เช่น กข 1234" />
            </Field>
          </div>

          {(() => {
            const subtotalBeforeVat = form.items.reduce((s, it) => s + lineTotal(it), 0);
            const vat = subtotalBeforeVat * ((Number(form.vatRate) || 0) / 100);
            const total = subtotalBeforeVat + vat;
            const totalQty = form.items.reduce((s, it) => s + (Number(it.qty) || 0), 0);
            const totalDeduct = form.items.reduce((s, it) => {
              const qty = Number(it.qty) || 0;
              const deductPct = Number(it.deductPct) || 0;
              const deductKg = Number(it.deductKg) || 0;
              return s + (qty * deductPct / 100) + deductKg;
            }, 0);
            const totalNet = totalQty - totalDeduct;
            return (
              <div style={{ marginTop: 16 }}>
                {/* สรุปน้ำหนัก */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 12 }}>
                  {[
                    { label: "น้ำหนักรวม", value: fmt(totalQty), unit: "กก.", color: "#1f2937", bg: "#f9fafb" },
                    { label: "รวมน้ำหนักหัก", value: fmt(totalDeduct), unit: "กก.", color: "#993c1d", bg: "#fef2f2" },
                    { label: "น้ำหนักสุทธิ", value: fmt(totalNet), unit: "กก.", color: "#0f6e56", bg: "#f0fdf4" },
                  ].map((item) => (
                    <div key={item.label} style={{ background: item.bg, borderRadius: 10, padding: "10px 14px", textAlign: "center", border: `1px solid ${item.color}22` }}>
                      <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>{item.label}</div>
                      <div style={{ fontSize: 22, fontWeight: 800, color: item.color }}>{item.value}</div>
                      <div style={{ fontSize: 12, color: item.color, fontWeight: 600 }}>{item.unit}</div>
                    </div>
                  ))}
                </div>
                {/* สรุปยอดเงิน */}
                <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", fontSize: 14 }}>
                  <Row label="ยอดก่อน VAT" value={`฿${fmt(subtotalBeforeVat)}`} />
                  {vat > 0 && <Row label={`VAT ${form.vatRate}%`} value={`+฿${fmt(vat)}`} color="#993c1d" />}
                  <Row label="ยอดรวมที่ต้องชำระ" value={`฿${fmt(total)}`} bold />
                  <p style={{ fontSize: 12, color: "#9ca3af", margin: "8px 0 0" }}>
                    * บันทึกใบนี้ก่อนได้เลย — ไปบันทึกการจ่ายเงินจริงที่เมนู "รับชำระ/จ่ายชำระ" ทีหลังได้
                  </p>
                </div>
              </div>
            );
          })()}

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button data-kbsubmit style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
          </div>{/* end data-kbform */}
        </Modal>
      )}

      {modal && modal.mode === "view" && (
        <PurchasePdfModal po={purchases.find((p) => p.id === modal.item.id) || modal.item} customer={customers.find((c) => c.id === modal.item.customerId)} products={products} storeBankAccounts={storeBankAccounts} companySettings={companySettings} onClose={() => setModal(null)} />
      )}
      </div>{/* end tab-export-purchases */}
    </div>
  );
}

function PurchasePdfModal({ po, customer, products, storeBankAccounts, companySettings, onClose }) {
  const cs = companySettings || {};
  const prodInfo = (id) => products.find((p) => p.id === id) || { name: id, unit: "" };

  const calcNet = (it) => {
    const qty = Number(it.qty) || 0;
    let net;
    // รองรับ field ใหม่ (deductPct/deductKg) ก่อน
    if (it.deductPct != null || it.deductKg != null) {
      const deductPct = Number(it.deductPct) || 0;
      const deductKg = Number(it.deductKg) || 0;
      net = qty - (qty * deductPct / 100) - deductKg;
    } else {
      const deduct = Number(it.deduct) || 0;
      if (it.deductType === "pct") net = qty * (1 - deduct / 100);
      else if (it.net != null) return Number(it.net); // ผู้ใช้กรอกเองไม่ round ทับ
      else net = qty - deduct;
    }
    return Math.round(net * 100) / 100; // จำกัดทศนิยม 2 ตำแหน่ง
  };

  const subtotal = po.items.reduce((s, it) => {
    const net = calcNet(it);
    const discountPct = Number(it.discountPct) || 0;
    return s + net * (Number(it.price) || 0) * (1 - discountPct / 100);
  }, 0);
  const vat = subtotal * ((Number(po.vatRate) || 0) / 100);
  const total = subtotal + vat;
  const primaryColor = cs.primaryColor || "#0f6e56";
  // style แถวเตี้ยเฉพาะใบรับสินค้า (ลดระยะห่างบน-ล่าง ไม่กระทบตารางหน้าอื่น)
  const thCompact = { ...thStyle, padding: "4px 12px" };
  const tdCompact = { ...tdStyle, padding: "4px 12px" };

  return (
    <Modal title={`${cs.purchaseTitle || "ใบรับสินค้า"} ${po.id}`} onClose={onClose} wide>
      <div id="purchase-pdf-content" style={{ background: "#fff", padding: "16px", border: "1px solid #e5e7eb", borderRadius: 8, fontFamily: "'Noto Sans Thai', sans-serif", fontSize: 12 }}>
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${primaryColor}`, paddingBottom: 12, marginBottom: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {cs.logo && (
              <img src={cs.logo} alt="logo" style={{ height: 60, maxWidth: 120, objectFit: "contain" }} />
            )}
            <div>
              <div style={{ fontSize: 18, fontWeight: 700, color: primaryColor }}>{cs.name || "วงจรกรีน รีไซเคิล"}</div>
              {cs.nameEn && <div style={{ fontSize: 12, color: "#6b7280" }}>{cs.nameEn}</div>}
              {cs.taxId && <div style={{ fontSize: 12, color: "#6b7280" }}>เลขผู้เสียภาษี: {cs.taxId}</div>}
              {cs.address && <div style={{ fontSize: 12, color: "#6b7280" }}>{cs.address}</div>}
              {cs.phone && <div style={{ fontSize: 12, color: "#6b7280" }}>โทร: {cs.phone}</div>}
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 16, fontWeight: 700, color: primaryColor }}>{cs.purchaseTitle || "ใบรับสินค้า"}</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>เลขที่: {po.id}</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>วันที่: {po.date}</div>
            <div style={{ fontSize: 12, marginTop: 4 }}>
              สถานะ: <span style={{ fontWeight: 600, color: po.status === "อนุมัติแล้ว" ? "#0f6e56" : po.status === "ยกเลิก" ? "#a32d2d" : "#854f0b" }}>{po.status || "รออนุมัติ"}</span>
            </div>
          </div>
        </div>

        <div style={{ marginBottom: 16, fontSize: 13 }}>
          <div style={{ fontWeight: 600, marginBottom: 4 }}>ข้อมูลผู้ขาย/ผู้ส่งสินค้า</div>
          <div>{customer?.name}</div>
          <div style={{ color: "#6b7280" }}>{customer?.address}</div>
          <div style={{ color: "#6b7280" }}>โทร: {customer?.phone} | เลขผู้เสียภาษี: {customer?.taxId}</div>
          {po.vehiclePlate && (
            <div style={{ marginTop: 4, color: "#374151" }}>
              🚛 ทะเบียนรถ: <strong>{po.vehiclePlate}</strong>
            </div>
          )}
        </div>

        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, tableLayout: "fixed" }}>
          <thead>
            <tr style={{ background: primaryColor + "22" }}>
              <th style={{ ...thCompact, color: primaryColor, width: "35%" }}>สินค้า</th>
              <th style={{ ...thCompact, color: primaryColor, textAlign: "right", width: "13%" }}>จำนวน</th>
              <th style={{ ...thCompact, color: primaryColor, textAlign: "right", width: "13%" }}>รวมหัก</th>
              <th style={{ ...thCompact, color: primaryColor, textAlign: "right", width: "13%" }}>สุทธิ</th>
              <th style={{ ...thCompact, color: primaryColor, textAlign: "right", width: "13%" }}>ราคา/หน่วย</th>
              <th style={{ ...thCompact, color: primaryColor, textAlign: "right", width: "13%" }}>จำนวนเงิน</th>
            </tr>
          </thead>
          <tbody>
            {po.items.map((it, idx) => {
              const p = prodInfo(it.productId);
              const qty = Number(it.qty) || 0;
              const net = calcNet(it);
              const deducted = qty - net;
              const discountPct = Number(it.discountPct) || 0;
              const amount = net * (Number(it.price) || 0) * (1 - discountPct / 100);
              return (
                <tr key={idx}>
                  <td style={{ ...tdCompact, wordBreak: "break-word" }}>{p.name}</td>
                  <td style={{ ...tdCompact, textAlign: "right" }}>{fmt(qty)} {p.unit}</td>
                  <td style={{ ...tdCompact, textAlign: "right", color: deducted > 0 ? "#993c1d" : "#9ca3af" }}>{deducted > 0 ? fmt(deducted) : "0"}</td>
                  <td style={{ ...tdCompact, textAlign: "right" }}>{fmt(net)}</td>
                  <td style={{ ...tdCompact, textAlign: "right" }}>{fmt(it.price)}</td>
                  <td style={{ ...tdCompact, textAlign: "right", fontWeight: 600 }}>{fmt(amount)}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            {(() => {
              const totalQty = po.items.reduce((s, it) => s + (Number(it.qty) || 0), 0);
              const totalNet = po.items.reduce((s, it) => {
                const qty = Number(it.qty) || 0;
                const net = it.deductType === "pct" ? qty*(1-(Number(it.deductPct)||0)/100) : qty - (Number(it.deduct)||0);
                return s + net;
              }, 0);
              const totalDeducted = totalQty - totalNet;
              const unit = po.items[0] ? (products.find(p=>p.id===po.items[0].productId)?.unit || "") : "";
              return (
                <tr style={{ background: "#f9fafb" }}>
                  <td style={{ ...tdCompact, fontWeight: 700, color: "#374151" }}>รวมทั้งหมด</td>
                  <td style={{ ...tdCompact, textAlign: "right", fontWeight: 700 }}>{fmt(totalQty)}</td>
                  <td style={{ ...tdCompact, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>{fmt(totalDeducted)}</td>
                  <td style={{ ...tdCompact, textAlign: "right", fontWeight: 700 }}>{fmt(totalNet)}</td>
                  <td style={{ ...tdCompact }}></td>
                  <td style={{ ...tdCompact }}></td>
                </tr>
              );
            })()}
            {po.vatRate > 0 && (
              <tr>
                <td colSpan={5} style={{ ...tdCompact, textAlign: "right", fontSize: 11 }}>ยอดก่อน VAT</td>
                <td style={{ ...tdCompact, textAlign: "right", fontSize: 11 }}>{fmt(subtotal)} บาท</td>
              </tr>
            )}
            {po.vatRate > 0 && (
              <tr>
                <td colSpan={5} style={{ ...tdCompact, textAlign: "right", fontSize: 11, color: "#993c1d" }}>VAT {po.vatRate}%</td>
                <td style={{ ...tdCompact, textAlign: "right", fontSize: 11, color: "#993c1d" }}>+{fmt(vat)} บาท</td>
              </tr>
            )}
            <tr style={{ background: "#f0fdf4" }}>
              <td colSpan={5} style={{ ...tdCompact, textAlign: "right", fontWeight: 700, fontSize: 13 }}>จำนวนเงินสุทธิ</td>
              <td style={{ ...tdCompact, textAlign: "right", fontWeight: 700, fontSize: 13, color: "#0f6e56" }}>{fmt(total)}</td>
            </tr>
          </tfoot>
        </table>

        <div style={{ pageBreakInside: "avoid", pageBreakBefore: "auto" }}>

          {cs.footerNote && (
            <div style={{ marginTop: 8, padding: "6px 10px", background: "#f9fafb", borderRadius: 6, fontSize: 11, color: "#6b7280" }}>
              {cs.footerNote}
            </div>
          )}

          {cs.showSignature !== false && (
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 24, fontSize: 11 }}>
              <div style={{ textAlign: "center", width: "45%" }}>
                <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้รับสินค้า</div>
              </div>
              <div style={{ textAlign: "center", width: "45%" }}>
                <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้ส่งสินค้า / ลูกค้า</div>
              </div>
            </div>
          )}

          {(po.payments || []).length > 0 && (
            <div style={{ marginTop: 16, borderTop: "1px dashed #d1d5db", paddingTop: 10 }}>
              <div style={{ fontWeight: 600, fontSize: 11, marginBottom: 4 }}>รายละเอียดช่องทางการชำระเงิน</div>
              <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 6 }}>
                ช่องทางชำระเงิน: {po.paymentMethod || "-"}
                {(() => {
                  const b = (customer?.bankAccounts || []).find((x) => x.id === po.receivingCustomerBankId);
                  return b ? ` — บัญชีรับเงิน: ${b.bankName} ${b.accountNo} (${b.accountName})` : "";
                })()}
              </div>
              <div style={{ textAlign: "right", fontSize: 11, fontWeight: 600 }}>
                ชำระแล้วทั้งหมด: {fmt((po.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0))} บาท
                {" / "}คงเหลือ: {fmt(total - (po.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0))} บาท
              </div>
            </div>
          )}
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <button style={btnSecondary} onClick={onClose}>ปิด</button>
        <button style={btnPrimary} onClick={() => printAsPDF("purchase-pdf-content", `${cs.purchaseTitle || "ใบรับสินค้า"} ${po.id}`)}><Download size={16} /> พิมพ์ / บันทึก PDF</button>
      </div>
    </Modal>
  );
}

// ===================================================================
// SALES TAB
// ===================================================================
// WITHDRAWALS TAB (เบิกสินค้าเพื่อขาย)
// ===================================================================
// ซิงค์ยอดเบิกที่ผูกกับใบขายแต่ละใบ: รวมจำนวน+มูลค่าของทุก LOT ตาม (targetSaleId, targetProductId)
// แล้วเขียนกลับเป็นรายการสินค้า (item) ในใบขายนั้น (สร้าง/อัปเดต/ลบ ตามความเหมาะสม)
// ราคาขาย (price) ที่ผู้ใช้กำหนดไว้เองจะยังคงอยู่ ไม่ถูกเขียนทับ
function syncWithdrawalsToSales(sales, withdrawalLots) {
  // group ทุกรายการเบิกในทุก LOT ตาม targetSaleId -> targetProductId
  const bySale = {};
  withdrawalLots.forEach((lot) => {
    if (!lot.targetSaleId) return;
    (lot.items || []).forEach((it) => {
      if (!it.targetProductId) return;
      if (!bySale[lot.targetSaleId]) bySale[lot.targetSaleId] = {};
      if (!bySale[lot.targetSaleId][it.targetProductId]) bySale[lot.targetSaleId][it.targetProductId] = { qty: 0, value: 0 };
      bySale[lot.targetSaleId][it.targetProductId].qty += Number(it.qty) || 0;
      bySale[lot.targetSaleId][it.targetProductId].value += Number(it.value) || 0;
    });
  });

  return sales.map((inv) => {
    const groups = bySale[inv.id];
    // เก็บรายการที่ไม่ใช่มาจากการเบิก ไว้เหมือนเดิม
    const nonWithdrawalItems = inv.items.filter((it) => !it.fromWithdrawal);
    if (!groups) {
      // ไม่มีรายการเบิกผูกกับใบนี้ -> เอารายการ fromWithdrawal เดิมออกถ้ามี (ไม่มีข้อมูลเบิกแล้ว)
      return inv.items.some((it) => it.fromWithdrawal) ? { ...inv, items: nonWithdrawalItems } : inv;
    }
    const withdrawalItems = Object.entries(groups).map(([targetProductId, g]) => {
      const existing = inv.items.find((it) => it.fromWithdrawal && it.productId === targetProductId);
      const avgCost = g.qty > 0 ? g.value / g.qty : 0;
      // คงค่า "จำนวนหัก" ที่ผู้ใช้แก้ไขเองไว้ (ถ้ามี) แล้วคำนวณจำนวนสุทธิใหม่จาก qty - deduct
      const deduct = existing ? (Number(existing.deduct) || 0) : 0;
      const net = g.qty - deduct;
      return {
        productId: targetProductId,
        qty: g.qty,
        deduct,
        net,
        price: existing ? existing.price : Math.round(avgCost * 100) / 100,
        fromWithdrawal: true,
        withdrawalCost: avgCost,
        withdrawalValue: g.value,
      };
    });
    return { ...inv, items: [...nonWithdrawalItems, ...withdrawalItems] };
  });
}

function WithdrawalsTab({ products, purchases, sales, setSales, withdrawals, setWithdrawals, inventory, customers, companySettings }) {
  const cs = companySettings || {};
  const [modal, setModal] = useState(null); // {mode:'add'|'edit'}
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [aggregateSearch, setAggregateSearch] = useState("");
  const [expandedGroups, setExpandedGroups] = useState({});
  const toggleGroup = (saleId) => setExpandedGroups((prev) => ({ ...prev, [saleId]: !prev[saleId] }));
  const [expanded, setExpanded] = useState(null);
  const [printLot, setPrintLot] = useState(null); // ใบเบิกสินค้าที่กำลังจะพิมพ์

  const prodName = (id) => products.find((p) => p.id === id)?.name || id;
  const prodUnit = (id) => products.find((p) => p.id === id)?.unit || "";
  const custName = (id) => customers.find((c) => c.id === id)?.name || "";

  const blankLineItem = () => ({ sourceProductId: "", qty: 0, targetProductId: "" });

  const blankForm = () => ({
    id: genId("WD", withdrawals, new Date().toISOString().slice(0, 10)),
    date: new Date().toISOString().slice(0, 10),
    targetSaleMode: "existing", // existing | new
    targetSaleId: sales[0]?.id || "",
    newSaleId: "",
    items: [blankLineItem()],
  });
  const [form, setForm] = useState(blankForm());

  // auto-fill newSaleId when switching to "สร้างใบขายใหม่"
  const setTargetSaleMode = (mode) => {
    if (mode === "new") {
      setForm((f) => ({ ...f, targetSaleMode: mode, newSaleId: genId("INV", sales) }));
    } else {
      setForm((f) => ({ ...f, targetSaleMode: mode, newSaleId: "" }));
    }
  };

  const openAdd = () => { setForm(blankForm()); setModal({ mode: "add" }); };
  const openEdit = (lot) => {
    setForm({
      ...JSON.parse(JSON.stringify(lot)),
      targetSaleMode: "existing",
      newSaleId: "",
      items: lot.items.map((it) => ({ sourceProductId: it.sourceProductId, qty: it.qty, targetProductId: it.targetProductId })),
    });
    setModal({ mode: "edit", item: lot });
  };

  // สต๊อกฐานสำหรับคำนวณ preview: กรณีแก้ไข ให้ตัด LOT เดิมออกก่อน เพื่อความถูกต้อง
  const baseInventory = useMemo(() => {
    if (modal && modal.mode === "edit") {
      const withoutThis = withdrawals.filter((w) => w.id !== modal.item.id);
      return computeInventory(products, purchases, sales, withoutThis);
    }
    return inventory;
  }, [inventory, modal, withdrawals, products, purchases, sales]);

  // preview ของแต่ละรายการในฟอร์ม: ต้องคำนวณทีละรายการตามลำดับ เพราะรายการเดียวกัน (สินค้าต้นทาง)
  // ที่ปรากฏหลายแถวใน LOT เดียวกัน ต้องตัดสต๊อกต่อเนื่องกัน ไม่ใช่คำนวณจากสต๊อกตั้งต้นซ้ำ
  const previews = useMemo(() => {
    // clone lots จาก baseInventory เพื่อจำลองการตัดสต๊อกต่อเนื่องในฟอร์มนี้
    const lotsClone = {};
    Object.keys(baseInventory.lots).forEach((pid) => {
      lotsClone[pid] = baseInventory.lots[pid].map((l) => ({ ...l }));
    });
    const fakeInventory = { ...baseInventory, lots: lotsClone };
    return form.items.map((it) => {
      const result = computeWithdrawalCost(fakeInventory, it.sourceProductId, Number(it.qty) || 0);
      // หักสต๊อกจำลองออกจริง เพื่อให้แถวถัดไปคำนวณต่อเนื่อง
      let remaining = Number(it.qty) || 0;
      const lots = fakeInventory.lots[it.sourceProductId] || [];
      for (let i = 0; i < lots.length && remaining > 0; i++) {
        const lot = lots[i];
        if (lot.qtyRemaining <= 0) continue;
        const take = Math.min(lot.qtyRemaining, remaining);
        lot.qtyRemaining -= take;
        remaining -= take;
      }
      return result;
    });
  }, [baseInventory, form.items]);

  const stockRemaining = useMemo(() => {
    // คงเหลือสต๊อกของสินค้าต้นทาง โดยพิจารณายอดที่ถูกใช้ไปแล้วจากแถวก่อนหน้าในฟอร์มเดียวกัน
    const result = {};
    const used = {};
    form.items.forEach((it, idx) => {
      const base = baseInventory.summary.find((s) => s.productId === it.sourceProductId)?.qty || 0;
      const usedSoFar = used[it.sourceProductId] || 0;
      result[idx] = base - usedSoFar;
      used[it.sourceProductId] = usedSoFar + (Number(it.qty) || 0);
    });
    return result;
  }, [baseInventory, form.items]);

  const updateLineItem = (idx, field, value) => {
    const items = [...form.items];
    items[idx] = { ...items[idx], [field]: value };
    setForm({ ...form, items });
  };
  const addLineItem = () => setForm({ ...form, items: [...form.items, blankLineItem()] });
  const removeLineItem = (idx) => setForm({ ...form, items: form.items.filter((_, i) => i !== idx) });

  const lotTotalValue = previews.reduce((s, p) => s + p.value, 0);

  const save = () => {
    let targetSaleId = form.targetSaleMode === "new" ? form.newSaleId.trim() : form.targetSaleId;
    if (!targetSaleId) return;
    if (form.items.length === 0) return;

    const lineItems = form.items.map((it, idx) => {
      const qty = Number(it.qty) || 0;
      const { value, shortfall } = previews[idx];
      return {
        sourceProductId: it.sourceProductId,
        qty,
        value,
        avgCost: qty > 0 ? value / qty : 0,
        shortfall,
        targetProductId: it.targetProductId,
      };
    }).filter((it) => it.qty > 0 && it.sourceProductId && it.targetProductId);

    if (lineItems.length === 0) return;

    const newLot = {
      id: form.id.trim() || blankForm().id,
      date: form.date,
      targetSaleId,
      items: lineItems,
    };

    let updatedSales = sales;
    // สร้างใบขายใหม่ถ้าจำเป็น
    if (!sales.find((s) => s.id === targetSaleId)) {
      updatedSales = [...sales, {
        id: targetSaleId, date: form.date, customerId: customers[0]?.id || "",
        items: [], discount: 0, vatRate: 7, paymentMethod: PAYMENT_METHODS[0], paymentStatus: PAYMENT_STATUSES[0],
      }];
    }

    const updatedWithdrawals = modal.mode === "edit"
      ? withdrawals.map((w) => (w.id === modal.item.id ? newLot : w))
      : [...withdrawals, newLot];

    setSales(syncWithdrawalsToSales(updatedSales, updatedWithdrawals));
    setWithdrawals(updatedWithdrawals);
    setModal(null);
  };

  const remove = (id) => {
    const updatedWithdrawals = withdrawals.filter((w) => w.id !== id);
    setSales(syncWithdrawalsToSales(sales, updatedWithdrawals));
    setWithdrawals(updatedWithdrawals);
  };

  const filtered = withdrawals.filter((w) =>
    w.id.includes(search) || (w.targetSaleId || "").includes(search) ||
    (w.items || []).some((it) => prodName(it.sourceProductId).includes(search) || prodName(it.targetProductId).includes(search))
  ).filter((w) => (!dateFrom || (w.date || "") >= dateFrom) && (!dateTo || (w.date || "") <= dateTo)).sort((a, b) => (b.date || "").localeCompare(a.date || "") || b.id.localeCompare(a.id));

  const lotTotal = (lot) => (lot.items || []).reduce((s, it) => s + (Number(it.value) || 0), 0);
  const lotQtyTotal = (lot) => (lot.items || []).reduce((s, it) => s + (Number(it.qty) || 0), 0);

  // สรุปยอดรวมของแต่ละใบขาย+สินค้าเป้าหมาย เพื่อแสดงตัวอย่างผลลัพธ์
  const aggregates = useMemo(() => {
    const groups = {};
    withdrawals.forEach((lot) => {
      if (!lot.targetSaleId) return;
      (lot.items || []).forEach((it) => {
        if (!it.targetProductId) return;
        const key = `${lot.targetSaleId}__${it.targetProductId}`;
        if (!groups[key]) groups[key] = { saleId: lot.targetSaleId, productId: it.targetProductId, qty: 0, value: 0 };
        groups[key].qty += Number(it.qty) || 0;
        groups[key].value += Number(it.value) || 0;
      });
    });
    return Object.values(groups).map((g) => ({ ...g, avgCost: g.qty > 0 ? g.value / g.qty : 0 }));
  }, [withdrawals]);

  const filteredAggregates = aggregates.filter((g) =>
    g.saleId.toLowerCase().includes(aggregateSearch.toLowerCase()) ||
    prodName(g.productId).toLowerCase().includes(aggregateSearch.toLowerCase())
  );

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 80px)" }}>
      <div style={{ flexShrink: 0 }}>
      <Header title="เบิกสินค้าเพื่อขาย" subtitle="เบิกสินค้าเป็น LOT (ตัดสต๊อกทันทีตามต้นทุน FIFO) เพื่อนำไปเปิดบิลขาย — 1 LOT เบิกได้หลายรายการ">
        <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> สร้างใบเบิกสินค้า (LOT)</button>
      </Header>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาเลขที่ LOT, สินค้า หรือเลข Invoice..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />
      </div>
      <div style={{ flex: 1, overflow: "auto" }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {filtered.map((lot) => {
          const isExpanded = expanded === lot.id;
          return (
            <div key={lot.id} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "14px 18px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
                <div style={{ minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                    <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 600, fontSize: 14, display: "inline-flex", alignItems: "center", gap: 6, color: "#534ab7" }}>
                      <PackageMinus size={14} /> {lot.id}
                    </span>
                    <span style={{ fontSize: 13, color: "#6b7280" }}>
                      → Invoice <span style={{ fontFamily: "'JetBrains Mono', monospace", fontWeight: 500 }}>{lot.targetSaleId}</span>
                      {sales.find((s) => s.id === lot.targetSaleId) && <> ({custName(sales.find((s) => s.id === lot.targetSaleId)?.customerId)})</>}
                    </span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 14, marginTop: 6, fontSize: 13, color: "#6b7280", flexWrap: "wrap" }}>
                    <span>วันที่: {lot.date}</span>
                    <span>{(lot.items || []).length} รายการเบิก</span>
                    <span>รวม {fmt(lotQtyTotal(lot))} หน่วย</span>
                  </div>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontSize: 12, color: "#9ca3af" }}>มูลค่ารวม</div>
                    <div style={{ fontSize: 18, fontWeight: 700, color: "#534ab7" }}>฿{fmt(lotTotal(lot))}</div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button style={iconBtn} onClick={() => setExpanded(isExpanded ? null : lot.id)} aria-label="รายละเอียด" title="ดูรายละเอียด">
                      {isExpanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    </button>
                    <button style={iconBtn} onClick={() => setPrintLot(lot)} aria-label="พิมพ์" title="พิมพ์ใบเบิกสินค้า"><Printer size={16} /></button>
                    <button style={iconBtn} onClick={() => openEdit(lot)} aria-label="แก้ไข"><Edit2 size={16} /></button>
                    <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบใบเบิกสินค้า "${lot.id}" ใช่หรือไม่?`, () => remove(lot.id))} aria-label="ลบ"><Trash2 size={16} /></button>
                  </div>
                </div>
              </div>

              {isExpanded && (
                <div style={{ marginTop: 14, paddingTop: 14, borderTop: "1px solid #f3f4f6" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr>
                        <th style={thStyle}>สินค้าที่เบิก (ต้นทาง)</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>จำนวนที่เบิก</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>มูลค่าที่เบิก</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ย/หน่วย</th>
                        <th style={thStyle}></th>
                        <th style={thStyle}>นำไปขายเป็นสินค้า (เป้าหมาย)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(lot.items || []).map((it, idx) => (
                        <tr key={idx}>
                          <td style={tdStyle}>{prodName(it.sourceProductId)}</td>
                          <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(it.qty)} {prodUnit(it.sourceProductId)}</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(it.value)}</td>
                          <td style={{ ...tdStyle, textAlign: "right" }}>
                            ฿{fmt(it.avgCost)}
                            {it.shortfall > 0 && <span style={{ color: "#a32d2d", fontSize: 11, marginLeft: 4 }}>(สต๊อกขาด {fmt(it.shortfall)})</span>}
                          </td>
                          <td style={{ ...tdStyle, color: "#9ca3af" }}><ArrowRight size={14} /></td>
                          <td style={tdStyle}>{prodName(it.targetProductId)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
        {filtered.length === 0 && (
          <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "24px", textAlign: "center", color: "#9ca3af" }}>
            ยังไม่มีรายการเบิกสินค้า
          </div>
        )}
      </div>

      {aggregates.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3 style={{ fontSize: 15, fontWeight: 600, margin: "0 0 10px" }}>สรุปยอดต้นทุนรวมที่ไปลงในใบขาย</h3>
          <SearchBar value={aggregateSearch} onChange={setAggregateSearch} placeholder="ค้นหาเลข Invoice หรือสินค้า..." />
          <Card>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={thStyle}>เลข Invoice</th>
                  <th style={thStyle}>สินค้าในใบขาย</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>จำนวนรวม</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>มูลค่ารวม</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ยใหม่</th>
                </tr>
              </thead>
              <tbody>
                {(() => {
                  const grouped = {};
                  filteredAggregates.forEach((g) => {
                    if (!grouped[g.saleId]) grouped[g.saleId] = [];
                    grouped[g.saleId].push(g);
                  });
                  if (Object.keys(grouped).length === 0) return (
                    <tr><td colSpan={5} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่พบรายการที่ค้นหา</td></tr>
                  );
                  return Object.entries(grouped).map(([saleId, items]) => {
                    const totalValue = items.reduce((s, g) => s + g.value, 0);
                    const totalQty = items.reduce((s, g) => s + g.qty, 0);
                    const isExpanded = !!expandedGroups[saleId];
                    return (
                      <React.Fragment key={saleId}>
                        {/* แถวกลุ่ม — คลิกเพื่อ expand/collapse */}
                        <tr
                          onClick={() => toggleGroup(saleId)}
                          style={{ background: "#f3f4f6", cursor: "pointer" }}
                          onMouseEnter={(e) => e.currentTarget.style.background = "#e5e7eb"}
                          onMouseLeave={(e) => e.currentTarget.style.background = "#f3f4f6"}
                        >
                          <td style={{ ...tdStyle, fontWeight: 700, color: "#534ab7", fontFamily: "'JetBrains Mono', monospace", display: "flex", alignItems: "center", gap: 6 }}>
                            {isExpanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                            {saleId}
                          </td>
                          <td style={{ ...tdStyle, color: "#6b7280", fontSize: 12 }}>{items.length} รายการ</td>
                          <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280" }}>{fmt(totalQty)}</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#534ab7" }}>฿{fmt(totalValue)}</td>
                          <td style={{ ...tdStyle, textAlign: "right", color: "#9ca3af" }}>—</td>
                        </tr>
                        {/* รายการสินค้า — แสดงเมื่อ expand */}
                        {isExpanded && items.map((g) => (
                          <tr key={`${g.saleId}__${g.productId}`} style={{ background: "#fafafa" }}>
                            <td style={{ ...tdStyle, color: "#9ca3af", paddingLeft: 32 }}>↳</td>
                            <td style={tdStyle}>{prodName(g.productId)}</td>
                            <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(g.qty)} {prodUnit(g.productId)}</td>
                            <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(g.value)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#534ab7" }}>฿{fmt(g.avgCost)}</td>
                          </tr>
                        ))}
                      </React.Fragment>
                    );
                  });
                })()}
              </tbody>
            </table>
          </Card>
        </div>
      )}

      {modal && (
        <Modal title={modal.mode === "add" ? "สร้างใบเบิกสินค้า (LOT)" : `แก้ไขใบเบิกสินค้า · ${form.id}`} onClose={() => setModal(null)} wide fullscreen>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1.6fr", gap: "0 16px" }}>
            <Field label="เลขที่ใบเบิก (LOT)">
              <input style={inputStyle} value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} />
            </Field>
            <Field label="วันที่เบิก">
              <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value, id: genId("WD", withdrawals, e.target.value) })} />
            </Field>
            <Field label="นำไปเปิดบิลขาย (Invoice)">
              <div style={{ display: "flex", gap: 8 }}>
                <select style={inputStyle} value={form.targetSaleMode} onChange={(e) => setTargetSaleMode(e.target.value)}>
                  <option value="existing">เลือกใบขายเดิม</option>
                  <option value="new">สร้างใบขายใหม่</option>
                </select>
                {form.targetSaleMode === "existing" ? (
                  <select style={inputStyle} value={form.targetSaleId} onChange={(e) => setForm({ ...form, targetSaleId: e.target.value })}>
                    <option value="">-- เลือก Invoice --</option>
                    {sales.map((s) => <option key={s.id} value={s.id}>{s.id} · {custName(s.customerId)}</option>)}
                  </select>
                ) : (
                  <input style={inputStyle} placeholder="เช่น INV-2606-005" value={form.newSaleId} onChange={(e) => setForm({ ...form, newSaleId: e.target.value })} />
                )}
              </div>
            </Field>
          </div>
          <p style={{ fontSize: 12, color: "#9ca3af", marginTop: -8 }}>
            * ทุกรายการในใบเบิกนี้จะถูกนำไปรวมกับยอดเดิมในใบขายเดียวกัน หากสินค้าเป้าหมายซ้ำกับใบเบิกอื่น ระบบจะรวมจำนวน/มูลค่า และคำนวณราคาเฉลี่ยใหม่ให้อัตโนมัติ
          </p>

          <div style={{ marginTop: 8, marginBottom: 8, fontWeight: 600, fontSize: 14 }}>รายการเบิกสินค้า</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
              <thead>
                <tr>
                  <th style={{ ...thStyle, width: "22%" }}>สินค้าที่เบิก (ต้นทาง)</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "10%" }}>จำนวนที่เบิก</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "11%" }}>คงเหลือสต๊อก</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "11%" }}>มูลค่าที่เบิก</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "11%" }}>ราคาเฉลี่ย/หน่วย</th>
                  <th style={{ ...thStyle, width: "3%" }}></th>
                  <th style={{ ...thStyle, width: "22%" }}>นำไปขายเป็นสินค้า (เป้าหมาย)</th>
                  <th style={{ ...thStyle, width: "4%" }}></th>
                </tr>
              </thead>
              <tbody>
                {form.items.map((it, idx) => {
                  const p = previews[idx] || { value: 0, shortfall: 0 };
                  const qty = Number(it.qty) || 0;
                  const avgCost = qty > 0 ? p.value / qty : 0;
                  const remain = stockRemaining[idx] ?? 0;
                  return (
                    <tr key={idx}>
                      <td style={tdStyle}>
                        <ProductSelect products={products} value={it.sourceProductId} onChange={(pid) => updateLineItem(idx, "sourceProductId", pid)} />
                      </td>
                      <td style={tdStyle}><input type="number" style={{ ...inputStyle, width: "100%", textAlign: "right" }} value={it.qty} onChange={(e) => updateLineItem(idx, "qty", e.target.value)} /></td>
                      <td style={{ ...tdStyle, textAlign: "right", color: remain < 0 ? "#a32d2d" : "#6b7280" }}>{fmt(remain)} {prodUnit(it.sourceProductId)}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#3c3489" }}>฿{fmt(p.value)}</td>
                      <td style={{ ...tdStyle, textAlign: "right" }}>
                        ฿{fmt(avgCost)}
                        {p.shortfall > 0 && <div style={{ color: "#a32d2d", fontSize: 11 }}>(ขาด {fmt(p.shortfall)})</div>}
                      </td>
                      <td style={{ ...tdStyle, color: "#9ca3af" }}><ArrowRight size={14} /></td>
                      <td style={tdStyle}>
                        <ProductSelect products={products} value={it.targetProductId} onChange={(pid) => updateLineItem(idx, "targetProductId", pid)} />
                      </td>
                      <td style={tdStyle}><button style={btnDanger} onClick={() => removeLineItem(idx)}><Trash2 size={14} /></button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 10, flexWrap: "wrap", gap: 10 }}>
            <button style={btnSecondary} onClick={addLineItem}><Plus size={14} /> เพิ่มรายการเบิก</button>
            <div style={{ fontSize: 16, fontWeight: 700 }}>
              มูลค่าที่เบิกรวม: <span style={{ color: "#3c3489" }}>฿{fmt(lotTotalValue)}</span>
            </div>
          </div>

          {/* สรุปน้ำหนัก */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginTop: 12 }}>
            {(() => {
              const totalQty = form.items.reduce((s, it) => s + (Number(it.qty) || 0), 0);
              const totalNet = form.items.reduce((s, it) => s + (Number(it.qty) || 0) - (Number(it.value) / (Number(it.price) || 1) || 0), 0);
              const totalDeduct = form.items.reduce((s, it) => s + (Number(it.containerWeight) || 0), 0);
              const netWeight = totalQty - totalDeduct;
              return [
                { label: "น้ำหนักรวม", value: fmt(totalQty), color: "#1f2937", bg: "#f9fafb" },
                { label: "รวมน้ำหนักหัก", value: fmt(totalDeduct), color: "#993c1d", bg: "#fef2f2" },
                { label: "น้ำหนักสุทธิ", value: fmt(netWeight), color: "#3c3489", bg: "#f5f3ff" },
              ].map((item) => (
                <div key={item.label} style={{ background: item.bg, borderRadius: 10, padding: "10px 14px", textAlign: "center", border: `1px solid ${item.color}22` }}>
                  <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>{item.label}</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: item.color }}>{item.value}</div>
                  <div style={{ fontSize: 12, color: item.color, fontWeight: 600 }}>กก.</div>
                </div>
              ));
            })()}
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {printLot && (
        <Modal title={`ใบเบิกสินค้า · ${printLot.id}`} onClose={() => setPrintLot(null)} wide>
          <div id="withdrawal-pdf-content" style={{ background: "#fff", padding: 24, border: "1px solid #e5e7eb", borderRadius: 8 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${cs.accentColor || "#3c3489"}`, paddingBottom: 12, marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                {cs.logo && (
                  <img src={cs.logo} alt="logo" style={{ height: 50, maxWidth: 100, objectFit: "contain" }} />
                )}
                <div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: cs.accentColor || "#3c3489" }}>{cs.name || "วงจรกรีน รีไซเคิล"}</div>
                  {cs.taxId && <div style={{ fontSize: 12, color: "#6b7280" }}>เลขผู้เสียภาษี: {cs.taxId}</div>}
                  {cs.address && <div style={{ fontSize: 12, color: "#6b7280" }}>{cs.address}</div>}
                  {cs.phone && <div style={{ fontSize: 12, color: "#6b7280" }}>โทร: {cs.phone}</div>}
                  <div style={{ fontSize: 12, color: "#6b7280" }}>ใบเบิกสินค้า (LOT)</div>
                </div>
              </div>
              <div style={{ textAlign: "right", fontSize: 12, color: "#6b7280" }}>
                <div>เลขที่: {printLot.id}</div>
                <div>วันที่: {printLot.date}</div>
                {printLot.targetSaleId && <div>อ้างอิงใบขาย: {printLot.targetSaleId}</div>}
                {sales.find((s) => s.id === printLot.targetSaleId) && (
                  <div>ลูกค้า: {custName(sales.find((s) => s.id === printLot.targetSaleId)?.customerId)}</div>
                )}
              </div>
            </div>

            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead><tr style={{ background: "#f3f4f6" }}>
                <th style={{ ...thStyle, padding: "4px 12px" }}>#</th>
                <th style={{ ...thStyle, padding: "4px 12px" }}>สินค้าต้นทาง</th>
                <th style={{ ...thStyle, padding: "4px 12px", textAlign: "right" }}>จำนวนที่เบิก</th>
                <th style={{ ...thStyle, padding: "4px 12px" }}>สินค้าปลายทาง</th>
                <th style={{ ...thStyle, padding: "4px 12px", textAlign: "right" }}>มูลค่า</th>
              </tr></thead>
              <tbody>
                {(printLot.items || []).map((it, i) => (
                  <tr key={i}>
                    <td style={{ ...tdStyle, padding: "4px 12px" }}>{i + 1}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px" }}>{prodName(it.sourceProductId)}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right" }}>{fmt(it.qty)} {prodUnit(it.sourceProductId)}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px" }}>{prodName(it.targetProductId)}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right" }}>฿{fmt(Number(it.value) || 0)}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={2} style={{ ...tdStyle, padding: "4px 12px", fontWeight: 700 }}>รวม</td>
                  <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right", fontWeight: 700 }}>{fmt(lotQtyTotal(printLot))}</td>
                  <td style={{ ...tdStyle, padding: "4px 12px" }}></td>
                  <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right", fontWeight: 700, color: "#3c3489" }}>฿{fmt(lotTotal(printLot))}</td>
                </tr>
              </tfoot>
            </table>

            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 48, fontSize: 12 }}>
              <div style={{ textAlign: "center", width: "45%" }}><div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้เบิกสินค้า</div></div>
              <div style={{ textAlign: "center", width: "45%" }}><div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้อนุมัติ</div></div>
            </div>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setPrintLot(null)}>ปิด</button>
            <button style={btnPrimary} onClick={() => printAsPDF("withdrawal-pdf-content", `ใบเบิกสินค้า ${printLot.id}`)}><Download size={16} /> พิมพ์ / บันทึก PDF</button>
          </div>
        </Modal>
      )}
    </div>
  </div>
  );
}
function SalesTab({ products, customers, sales, setSales, inventory, withdrawals, storeBankAccounts, companySettings }) {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");

  const blankItem = () => ({ productId: "", qty: 0, deduct: 0, price: 0 });
  const blankPayment = () => ({ id: "SP" + Date.now().toString().slice(-6), date: new Date().toISOString().slice(0, 10), amount: 0, method: PAYMENT_METHODS[0], toStoreBankId: "", note: "" });
  const blankForm = () => ({
    id: "", date: new Date().toISOString().slice(0, 10), customerId: "",
    items: [blankItem()], discount: 0, vatRate: 7, paymentStatus: PAYMENT_STATUSES[0],
    payments: [], vehiclePlate: "",
  });
  const [form, setForm] = useState(blankForm());

  const custName = (id) => customers.find((c) => c.id === id)?.name || id;
  const prodName = (id) => products.find((p) => p.id === id)?.name || id;
  const prodUnit = (id) => products.find((p) => p.id === id)?.unit || "";

  const filtered = sales.filter((inv) => inv.id.includes(search) || custName(inv.customerId).includes(search)).filter((inv) => (!dateFrom || (inv.date || "") >= dateFrom) && (!dateTo || (inv.date || "") <= dateTo)).sort((a, b) => (b.date || "").localeCompare(a.date || "") || b.id.localeCompare(a.id));

  const openAdd = () => { const _d2 = new Date().toISOString().slice(0, 10); setForm({ ...blankForm(), id: genId("INV", sales, _d2) }); setModal({ mode: "add" }); };
  const openEdit = (item) => {
    let payments = item.payments && item.payments.length > 0 ? [...item.payments] : [];
    setForm(JSON.parse(JSON.stringify({ ...item, payments })));
    setModal({ mode: "edit", item });
  };
  const openView = (item) => setModal({ mode: "view", item });

  const updateItem = (idx, field, value) => {
    const items = [...form.items];
    items[idx] = { ...items[idx], [field]: value };
    setForm({ ...form, items });
  };
  const addItem = () => setForm({ ...form, items: [...form.items, blankItem()] });
  const removeItem = (idx) => setForm({ ...form, items: form.items.filter((_, i) => i !== idx) });

  const addPayment = () => setForm({ ...form, payments: [...(form.payments || []), blankPayment()] });
  const updatePayment = (idx, field, value) => {
    const payments = [...(form.payments || [])];
    payments[idx] = { ...payments[idx], [field]: value };
    setForm({ ...form, payments });
  };
  const removePayment = (idx) => setForm({ ...form, payments: (form.payments || []).filter((_, i) => i !== idx) });

  const lineNet = (it) => (Number(it.qty) || 0) - (Number(it.deduct) || 0);
  const lineTotal = (it) => lineNet(it) * (Number(it.price) || 0);
  const subtotal = form.items.reduce((s, it) => s + lineTotal(it), 0);
  const afterDiscount = subtotal - (Number(form.discount) || 0);
  const vatAmount = afterDiscount * ((Number(form.vatRate) || 0) / 100);
  const grandTotal = afterDiscount + vatAmount;
  const cogs = form.items.reduce((s, it) => s + (it.fromWithdrawal ? (it.withdrawalValue || 0) : 0), 0);
  const profit = afterDiscount - cogs;
  const totalPaid = (form.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const remaining = grandTotal - totalPaid;

  const calcInvoiceTotals = (inv) => {
    const sub = inv.items.reduce((s, it) => s + (it.deductType === "pct" ? (Number(it.qty)||0)*(1-(Number(it.deduct)||0)/100) : (it.net != null ? Number(it.net) : (Number(it.qty)||0)-(Number(it.deduct)||0))) * (Number(it.price)||0) * (1-(Number(it.discountPct)||0)/100), 0);
    const ad = sub - (inv.discount || 0);
    const vat = ad * ((inv.vatRate || 0) / 100);
    const total = ad + vat;
    const paid = (inv.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
    return { sub, ad, vat, total, paid, remaining: total - paid };
  };

  const save = () => {
    if (!form.id.trim() || form.items.length === 0) return;
    const cleaned = {
      ...form,
      discount: Number(form.discount) || 0,
      vatRate: Number(form.vatRate) || 0,
      items: form.items.map((it) => ({ ...it, qty: Number(it.qty) || 0, deduct: Number(it.deduct) || 0, net: (Number(it.qty) || 0) - (Number(it.deduct) || 0), price: Number(it.price) || 0 })),
    };
    if (modal.mode === "add") setSales([...sales, cleaned]);
    else setSales(sales.map((s) => (s.id === modal.item.id ? cleaned : s)));
    setModal(null);
  };

  const remove = (id) => setSales(sales.filter((s) => s.id !== id));

  const statusColor = (st) => {
    if (st === "ชำระแล้ว") return { bg: "#eaf3de", color: "#27500a" };
    if (st === "ชำระบางส่วน") return { bg: "#faeeda", color: "#854f0b" };
    return { bg: "#fcebeb", color: "#791f1f" };
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "calc(100vh - 80px)" }}>
      <div style={{ flexShrink: 0 }}>
      <Header title="ระบบขายสินค้า (Sales)" subtitle="ออกใบ Invoice และบันทึกการขายสินค้ารีไซเคิล">
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <ExportToolbar
            onPDF={() => printAsPDF("tab-export-sales", "ขายสินค้า")}
            onExcel={() => {
              const rows = [
                ["เลข Invoice", "วันที่", "ลูกค้า", "สถานะ", "สินค้า", "จำนวนสุทธิ", "ราคา/หน่วย", "รวม"],
                ...filtered.flatMap((inv) =>
                  inv.items.map((it, i) => [
                    i === 0 ? inv.id : "", i === 0 ? inv.date : "", i === 0 ? custName(inv.customerId) : "", i === 0 ? inv.paymentStatus : "",
                    prodName(it.productId), it.net, it.price, it.net * it.price,
                  ])
                ),
              ];
              exportExcel(rows, "ขายสินค้า.xlsx", "ขายสินค้า");
            }}
            onImage={() => printAsPDF("tab-export-sales", "ขายสินค้า")}
          />
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> สร้าง Invoice</button>
        </div>
      </Header>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาเลข Invoice หรือชื่อลูกค้า..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />
      </div>
      <div id="tab-export-sales" style={{ flex: 1, overflow: "auto" }}>
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 860, tableLayout: "fixed" }}>
          <thead>
            <tr>
              <th style={{ ...thStyle, width: "15%" }}>เลข Invoice</th>
              <th style={{ ...thStyle, width: "10%" }}>วันที่</th>
              <th style={{ ...thStyle, width: "18%" }}>ลูกค้า</th>
              <th style={{ ...thStyle, width: "10%" }}>ทะเบียนรถ</th>
              <th style={{ ...thStyle, textAlign: "right", width: "13%" }}>ยอดสุทธิ</th>
              <th style={{ ...thStyle, textAlign: "right", width: "15%" }}>ยอดรับชำระ</th>
              <th style={{ ...thStyle, width: "10%" }}>สถานะ</th>
              <th style={{ ...thStyle, textAlign: "right", width: "9%" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((inv) => {
              const t = calcInvoiceTotals(inv);
              const livePayStatus = (inv.writeOff || t.remaining <= 0.01) ? "ชำระแล้ว" : t.paid > 0.01 ? "ชำระบางส่วน" : "ค้างรับ";
              const sc = statusColor(livePayStatus);
              return (
                <tr key={inv.id}>
                  <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", fontWeight: 500 }}>{inv.id}</td>
                  <td style={tdStyle}>{inv.date}</td>
                  <td style={tdStyle}>{custName(inv.customerId)}</td>
                  <td style={tdStyle}>
                    {inv.vehiclePlate ? <span style={{ background: "#f3f4f6", padding: "2px 8px", borderRadius: 4, fontSize: 12, fontFamily: "monospace" }}>🚛 {inv.vehiclePlate}</span> : <span style={{ color: "#d1d5db" }}>—</span>}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>{fmt(t.total)} บาท</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    <div style={{ fontSize: 12, color: "#0f6e56" }}>รับแล้ว ฿{fmt(t.paid)}</div>
                    {livePayStatus !== "ชำระแล้ว" && t.remaining > 0.01 && <div style={{ fontSize: 12, color: "#993c1d" }}>ค้าง ฿{fmt(t.remaining)}</div>}
                  </td>
                  <td style={tdStyle}><span style={{ background: sc.bg, color: sc.color, padding: "2px 10px", borderRadius: 6, fontSize: 12, fontWeight: 500 }}>{livePayStatus}</span></td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    <div style={{ display: "flex", gap: 4, justifyContent: "flex-end", flexWrap: "nowrap" }}>
                      <button style={{ ...iconBtn, padding: "4px 8px" }} title="ดู Invoice" onClick={() => openView(inv)}><Printer size={14} /></button>
                      <button style={{ ...iconBtn, padding: "4px 8px" }} title="แก้ไข" onClick={() => openEdit(inv)}><Edit2 size={14} /></button>
                      <button style={{ ...btnDanger, padding: "4px 8px" }} title="ลบ" onClick={() => confirmAction(`ต้องการลบใบขาย "${inv.id}" ใช่หรือไม่?`, () => remove(inv.id))}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่พบใบขายสินค้า</td></tr>}
          </tbody>
        </table>
        </div>
      </div>
      {modal && (modal.mode === "add" || modal.mode === "edit") && (
        <Modal title={modal.mode === "add" ? "สร้าง Invoice" : "แก้ไข Invoice"} onClose={() => setModal(null)} wide fullscreen>
          <div data-kbform>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 16px" }}>
            <Field label="เลข Invoice (กำหนดเอง)">
              <input style={inputStyle} value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)} placeholder="เช่น INV-2606-002" />
            </Field>
            <Field label="วันที่">
              <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value, id: genId("INV", sales, e.target.value) })} onKeyDown={(e) => handleEnterNavigate(e, save)} />
            </Field>
            <Field label="ลูกค้า">
              <CustomerSelect customers={customers} value={form.customerId} onChange={(cid) => setForm({ ...form, customerId: cid })} labelWithId={false} />
            </Field>
          </div>

          <div style={{ marginTop: 8, marginBottom: 8, fontWeight: 600, fontSize: 14 }}>รายการสินค้า</div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
              <thead>
                <tr>
                  <th style={{ ...thStyle, width: "28%" }}>สินค้า</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "9%" }}>จำนวน</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "9%" }}>จำนวนสุทธิ</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "9%" }}>จำนวนหัก</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "10%" }}>ราคา/หน่วย</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "10%" }}>ต้นทุนเฉลี่ย</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "12%" }}>คงเหลือสต๊อก</th>
                  <th style={{ ...thStyle, textAlign: "right", width: "9%" }}>รวม</th>
                  <th style={{ ...thStyle, width: "4%" }}></th>
                </tr>
              </thead>
              <tbody>
                {form.items.map((it, idx) => {
                  const stock = inventory.summary.find((s) => s.productId === it.productId);
                  const net = lineNet(it);
                  const fromW = !!it.fromWithdrawal;
                  const insufficient = !fromW && stock && net > stock.qty;
                  return (
                    <tr key={idx} style={fromW ? { background: "#eeedfe" } : undefined}>
                      <td style={tdStyle}>
                        <ProductSelect products={products} value={it.productId} onChange={(pid) => updateItem(idx, "productId", pid)} disabled={fromW} labelWithId={false} />
                        {fromW && (
                          <div style={{ fontSize: 11, color: "#534ab7", marginTop: 2, display: "flex", alignItems: "center", gap: 4 }}>
                            <PackageMinus size={12} /> มาจากการเบิกสินค้า (ตัดสต๊อกที่ใบเบิกแล้ว)
                          </div>
                        )}
                      </td>
                      <td style={tdStyle}>
                        {fromW ? (
                          <div style={{ textAlign: "right", color: "#534ab7", fontWeight: 500 }}>{fmt(it.qty)}</div>
                        ) : (
                          <input type="number" style={{ ...inputStyle, width: "100%", textAlign: "right" }} value={it.qty} onChange={(e) => updateItem(idx, "qty", e.target.value)} onKeyDown={(e) => handleEnterNavigate(e, save)} />
                        )}
                      </td>
                      <td style={tdStyle}>
                        <input
                          type="number"
                          style={{ ...inputStyle, width: "100%", textAlign: "right" }}
                          value={it.net != null ? it.net : net}
                          onChange={(e) => {
                            const newNet = e.target.value;
                            const newDeduct = (Number(it.qty) || 0) - (Number(newNet) || 0);
                            const items = [...form.items];
                            items[idx] = { ...items[idx], net: newNet, deduct: newDeduct };
                            setForm({ ...form, items });
                          }}
                          onKeyDown={(e) => handleEnterNavigate(e, save)}
                        />
                      </td>
                      <td style={tdStyle}>
                        <input
                          type="number"
                          style={{ ...inputStyle, width: "100%", textAlign: "right" }}
                          value={it.deduct}
                          onChange={(e) => {
                            const newDeduct = e.target.value;
                            const newNet = (Number(it.qty) || 0) - (Number(newDeduct) || 0);
                            const items = [...form.items];
                            items[idx] = { ...items[idx], deduct: newDeduct, net: newNet };
                            setForm({ ...form, items });
                          }}
                          onKeyDown={(e) => handleEnterNavigate(e, save)}
                        />
                      </td>
                      <td style={tdStyle}><input type="number" style={{ ...inputStyle, width: "100%", textAlign: "right" }} value={it.price} onChange={(e) => updateItem(idx, "price", e.target.value)} onKeyDown={(e) => handleEnterNavigate(e, save)} /></td>
                      <td style={{ ...tdStyle, textAlign: "right", color: fromW ? "#534ab7" : "#9ca3af", fontWeight: fromW ? 600 : 400 }}>
                        {fromW ? fmt(it.withdrawalCost || 0) : "—"}
                      </td>
                      <td style={{ ...tdStyle, textAlign: "right", color: insufficient ? "#a32d2d" : "#6b7280" }}>
                        {fromW ? <span style={{ color: "#9ca3af" }}>—</span> : <>{stock ? fmt(stock.qty) : "-"} {prodUnit(it.productId)}</>}
                      </td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>{fmt(lineTotal(it))}</td>
                      <td style={tdStyle}><button style={btnDanger} onClick={() => removeItem(idx)}><Trash2 size={14} /></button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <button style={{ ...btnSecondary, marginTop: 8 }} onClick={addItem}><Plus size={14} /> เพิ่มรายการสินค้า</button>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 16px", marginTop: 16 }}>
            <Field label="ส่วนลด (บาท)">
              <input type="number" style={inputStyle} value={form.discount} onChange={(e) => setForm({ ...form, discount: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)} />
            </Field>
            <Field label="VAT (%)">
              <input type="number" style={inputStyle} value={form.vatRate} onChange={(e) => setForm({ ...form, vatRate: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)} />
            </Field>
            <Field label="สถานะชำระเงิน">
              <select style={inputStyle} value={form.paymentStatus} onChange={(e) => setForm({ ...form, paymentStatus: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)}>
                {PAYMENT_STATUSES.map((m) => <option key={m} value={m}>{m}</option>)}
              </select>
            </Field>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginTop: 12 }}>
            {(() => {
              const totalQty = form.items.reduce((s, it) => s + (Number(it.qty) || 0), 0);
              const totalNet = form.items.reduce((s, it) => s + lineNet(it), 0);
              const totalDeduct = totalQty - totalNet;
              return [
                { label: "น้ำหนักรวม", value: fmt(totalQty), color: "#1f2937", bg: "#f9fafb" },
                { label: "รวมน้ำหนักหัก", value: fmt(totalDeduct), color: "#993c1d", bg: "#fef2f2" },
                { label: "น้ำหนักสุทธิ", value: fmt(totalNet), color: "#0f6e56", bg: "#f0fdf4" },
              ].map((item) => (
                <div key={item.label} style={{ background: item.bg, borderRadius: 10, padding: "10px 14px", textAlign: "center", border: `1px solid ${item.color}22` }}>
                  <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>{item.label}</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: item.color }}>{item.value}</div>
                  <div style={{ fontSize: 12, color: item.color, fontWeight: 600 }}>กก.</div>
                </div>
              ));
            })()}
          </div>

          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", marginTop: 8, fontSize: 14 }}>
            <Row label="ราคารวม (ก่อนหักส่วนลด)" value={`${fmt(subtotal)} บาท`} />
            <Row label="หลังหักส่วนลด" value={`${fmt(afterDiscount)} บาท`} />
            <Row label={`VAT (${form.vatRate}%)`} value={`${fmt(vatAmount)} บาท`} />
            <Row label="ยอดสุทธิ" value={`${fmt(grandTotal)} บาท`} bold />
            <div style={{ borderTop: "1px solid #e5e7eb", marginTop: 8, paddingTop: 8 }} />
            <Row label="ต้นทุนสินค้า (FIFO เฉลี่ย)" value={`${fmt(cogs)} บาท`} />
            <Row label="กำไรขั้นต้นโดยประมาณ" value={`${fmt(profit)} บาท`} bold color={profit >= 0 ? "#27500a" : "#791f1f"} />
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px", marginTop: 12 }}>
            <Field label="ทะเบียนรถ (ถ้ามี)">
              <input style={inputStyle} value={form.vehiclePlate || ""} onChange={(e) => setForm({ ...form, vehiclePlate: e.target.value })} onKeyDown={(e) => handleEnterNavigate(e, save)} placeholder="เช่น กข 1234" />
            </Field>
          </div>

          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", marginTop: 16, fontSize: 14 }}>
            <Row label="ยอดรวมที่ต้องเรียกเก็บ" value={`฿${fmt(grandTotal)}`} bold />
            <p style={{ fontSize: 12, color: "#9ca3af", margin: "8px 0 0" }}>
              * บันทึกใบนี้ก่อนได้เลย — ไปบันทึกการรับเงินจริงที่เมนู "รับชำระ/จ่ายชำระ" ทีหลังได้
            </p>
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button data-kbsubmit style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
          </div>{/* end data-kbform */}
        </Modal>
      )}

      {modal && modal.mode === "view" && (
        <SalesInvoiceModal inv={modal.item} customer={customers.find((c) => c.id === modal.item.customerId)} products={products} storeBankAccounts={storeBankAccounts} companySettings={companySettings} onClose={() => setModal(null)} />
      )}
    </div>
  );
}

function Row({ label, value, bold, color }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", padding: "3px 0", fontWeight: bold ? 700 : 400, fontSize: bold ? 15 : 13, color: color || "inherit" }}>
      <span>{label}</span>
      <span>{value}</span>
    </div>
  );
}

function SalesInvoiceModal({ inv, customer, products, storeBankAccounts, companySettings, onClose }) {
  const cs = companySettings || {};
  const prodInfo = (id) => products.find((p) => p.id === id) || { name: id, unit: "" };
  const subtotal = inv.items.reduce((s, it) => s + (it.deductType === "pct" ? (Number(it.qty)||0)*(1-(Number(it.deduct)||0)/100) : (it.net != null ? Number(it.net) : (Number(it.qty)||0)-(Number(it.deduct)||0))) * (Number(it.price)||0) * (1-(Number(it.discountPct)||0)/100), 0);
  const afterDiscount = subtotal - (inv.discount || 0);
  const vat = afterDiscount * ((inv.vatRate || 0) / 100);
  const total = afterDiscount + vat;
  const paid = (inv.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const remaining = total - paid;
  const accentColor = cs.accentColor || "#185fa5";

  return (
    <Modal title={`${cs.salesTitle || "Invoice"} ${inv.id}`} onClose={onClose} wide>
      <div id="sales-invoice-pdf-content" style={{ background: "#fff", padding: "24px", border: "1px solid #e5e7eb", borderRadius: 8 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${accentColor}`, paddingBottom: 12, marginBottom: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {cs.logo && (
              <img src={cs.logo} alt="logo" style={{ height: 60, maxWidth: 120, objectFit: "contain" }} />
            )}
            <div>
              <div style={{ fontSize: 18, fontWeight: 700, color: accentColor }}>{cs.name || "วงจรกรีน รีไซเคิล"}</div>
              {cs.nameEn && <div style={{ fontSize: 12, color: "#6b7280" }}>{cs.nameEn}</div>}
              {cs.taxId && <div style={{ fontSize: 12, color: "#6b7280" }}>เลขผู้เสียภาษี: {cs.taxId}</div>}
              {cs.address && <div style={{ fontSize: 12, color: "#6b7280" }}>{cs.address}</div>}
              {cs.phone && <div style={{ fontSize: 12, color: "#6b7280" }}>โทร: {cs.phone}</div>}
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 16, fontWeight: 700, color: accentColor }}>{cs.salesTitle || "ใบแจ้งหนี้ / Invoice"}</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>เลขที่: {inv.id}</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>วันที่: {inv.date}</div>
          </div>
        </div>

        <div style={{ marginBottom: 16, fontSize: 13 }}>
          <div style={{ fontWeight: 600, marginBottom: 4 }}>ลูกค้า</div>
          <div>{customer?.name}</div>
          <div style={{ color: "#6b7280" }}>{customer?.address}</div>
          <div style={{ color: "#6b7280" }}>โทร: {customer?.phone} | เลขผู้เสียภาษี: {customer?.taxId}</div>
          {inv.vehiclePlate && (
            <div style={{ marginTop: 4, color: "#374151" }}>
              🚛 ทะเบียนรถ: <strong>{inv.vehiclePlate}</strong>
            </div>
          )}
        </div>

        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
          <thead>
            <tr style={{ background: accentColor + "22" }}>
              <th style={{ ...thStyle, color: accentColor, padding: "4px 12px" }}>สินค้า</th>
              <th style={{ ...thStyle, color: accentColor, textAlign: "right", padding: "4px 12px" }}>น้ำหนักสุทธิ</th>
              <th style={{ ...thStyle, color: accentColor, textAlign: "right", padding: "4px 12px" }}>ราคา/หน่วย</th>
              <th style={{ ...thStyle, color: accentColor, textAlign: "right", padding: "4px 12px" }}>จำนวนเงิน</th>
            </tr>
          </thead>
          <tbody>
            {inv.items.map((it, idx) => {
              const p = prodInfo(it.productId);
              const net = it.deductType === "pct" ? (Number(it.qty)||0)*(1-(Number(it.deduct)||0)/100) : (it.net != null ? Number(it.net) : (Number(it.qty)||0)-(Number(it.deduct)||0));
              return (
                <tr key={idx}>
                  <td style={{ ...tdStyle, padding: "4px 12px" }}>{p.name}</td>
                  <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right" }}>{fmt(net)} {p.unit}</td>
                  <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right" }}>{fmt(it.price)}</td>
                  <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right", fontWeight: 600 }}>{fmt(net * it.price)}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            {(() => {
              const totalQty = inv.items.reduce((s, it) => s + (Number(it.qty)||0), 0);
              const totalNet = inv.items.reduce((s, it) => {
                const qty = Number(it.qty)||0;
                const net = it.deductType === "pct" ? qty*(1-(Number(it.deduct)||0)/100) : (it.net != null ? Number(it.net) : qty-(Number(it.deduct)||0));
                return s + net;
              }, 0);
              return (
                <tr style={{ background: "#f9fafb" }}>
                  <td style={{ ...tdStyle, padding: "4px 12px", fontWeight: 700 }}>รวมทั้งหมด</td>
                  <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right", fontWeight: 700 }}>{fmt(totalNet)}</td>
                  <td style={{ ...tdStyle, padding: "4px 12px" }}></td>
                  <td style={{ ...tdStyle, padding: "4px 12px" }}></td>
                </tr>
              );
            })()}
          </tfoot>
        </table>

        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 12 }}>
          <div style={{ width: 260 }}>
            <Row label="รวมเป็นเงิน" value={`${fmt(subtotal)} บาท`} />
            <Row label="ส่วนลด" value={`${fmt(inv.discount || 0)} บาท`} />
            <Row label="หลังหักส่วนลด" value={`${fmt(afterDiscount)} บาท`} />
            <Row label={`VAT ${inv.vatRate || 0}%`} value={`${fmt(vat)} บาท`} />
            <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderTop: "2px solid #185fa5", fontWeight: 700, fontSize: 15 }}>
              <span>ยอดสุทธิ</span>
              <span>{fmt(total)} บาท</span>
            </div>
          </div>
        </div>

        <div style={{ marginTop: 8, fontSize: 13 }}>
          สถานะ: <strong>{inv.paymentStatus}</strong>
        </div>

        {/* ลายเซ็น */}
        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 48, fontSize: 12 }}>
          <div style={{ textAlign: "center", width: "30%" }}>
            <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้ขาย / ผู้ส่งสินค้า</div>
          </div>
          <div style={{ textAlign: "center", width: "30%" }}>
            <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้รับสินค้า</div>
          </div>
          <div style={{ textAlign: "center", width: "30%" }}>
            <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้อนุมัติ</div>
          </div>
        </div>

        {/* รายละเอียดช่องทางการชำระเงิน */}
        {(inv.payments && inv.payments.length > 0) ? (
          <div style={{ marginTop: 24, borderTop: "1px dashed #d1d5db", paddingTop: 14 }}>
            <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8 }}>รายละเอียดช่องทางการชำระเงิน</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr>
                  <th style={thStyle}>วันที่รับ</th>
                  <th style={thStyle}>ช่องทาง</th>
                  <th style={thStyle}>บัญชีรับเงิน</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
                </tr>
              </thead>
              <tbody>
                {inv.payments.map((p, i) => {
                  const b = (storeBankAccounts || []).find((b) => b.id === p.toStoreBankId);
                  return (
                    <tr key={p.id || i}>
                      <td style={tdStyle}>{p.date}</td>
                      <td style={tdStyle}>{p.method}</td>
                      <td style={tdStyle}>{b ? `${b.bankName} ${b.accountNo}` : "-"}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#0f6e56" }}>฿{fmt(p.amount)}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{ borderTop: "2px solid #e5e7eb" }}>
                  <td colSpan={3} style={{ ...tdStyle, fontWeight: 700 }}>รับชำระแล้ว</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>฿{fmt(paid)}</td>
                </tr>
                {remaining > 0 && (
                  <tr>
                    <td colSpan={3} style={{ ...tdStyle, fontWeight: 700 }}>ยอดค้างชำระ</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>฿{fmt(remaining)}</td>
                  </tr>
                )}
              </tfoot>
            </table>
          </div>
        ) : (
          <div style={{ marginTop: 24, borderTop: "1px dashed #d1d5db", paddingTop: 10, fontSize: 13, color: "#9ca3af" }}>
            ยังไม่มีรายการรับชำระเงิน
          </div>
        )}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <button style={btnSecondary} onClick={onClose}>ปิด</button>
        <button style={btnPrimary} onClick={() => printAsPDF("sales-invoice-pdf-content", `${cs.salesTitle || "Invoice"} ${inv.id}`)}><Download size={16} /> พิมพ์ / บันทึก PDF</button>
      </div>
    </Modal>
  );
}

// ===================================================================
// PAYMENTS TAB (รับชำระ/จ่ายชำระ — รวมรายการค้างชำระจากใบรับสินค้าและใบขาย)
// ===================================================================
function PaymentsTab({ purchases, setPurchases, sales, setSales, customers, storeBankAccounts, deposits, expenses, setExpenses, companySettings, setCompanySettings, bankTransfers }) {
  const [showCreditSetting, setShowCreditSetting] = React.useState(false);
  const [creditDate, setCreditDate] = React.useState(new Date().toISOString().slice(0, 10));
  const [creditManual, setCreditManual] = React.useState(0); // ยอดตกหล่น กรอกมือ
  const [returnBankName, setReturnBankName] = React.useState(""); // ธนาคารโอนคืน
  const [returnBankNo, setReturnBankNo] = React.useState("");   // เลขที่บัญชีโอนคืน
  const [returnBankOwner, setReturnBankOwner] = React.useState(""); // ชื่อบัญชีโอนคืน
  const creditLimit = Number(companySettings?.creditLimit) || 0;
  const creditAccounts = companySettings?.creditAccounts || []; // array of storeBankAccount ids ที่นับในวงเงิน

  const [payFlags, setPayFlags] = React.useState(() => {
    try { return JSON.parse(localStorage.getItem("payFlags") || "{}"); } catch { return {}; }
  });
  const [showTransferSheet, setShowTransferSheet] = React.useState(false);
  const [transferTab, setTransferTab] = React.useState("purchase"); // "purchase" | "expense"
  const [transferDetailModal, setTransferDetailModal] = React.useState(null); // { row }
  const [transferEntries, setTransferEntries] = React.useState([{ bankId: "", amount: "" }]);
  // เก็บข้อมูลตั้งโอนแต่ละบิล { [id]: [{ bankId, amount }] }
  const [transferDetails, setTransferDetails] = React.useState(() => {
    try { return JSON.parse(localStorage.getItem("transferDetails") || "{}"); } catch { return {}; }
  });
  const saveTransferDetail = (id, entries) => {
    const next = { ...transferDetails, [id]: entries.filter(e => e.bankId || e.amount) };
    setTransferDetails(next);
    try { localStorage.setItem("transferDetails", JSON.stringify(next)); } catch {}
  };
  const setFlag = (id, flag, val) => {
    const next = { ...payFlags, [`${id}_${flag}`]: val };
    setPayFlags(next);
    try { localStorage.setItem("payFlags", JSON.stringify(next)); } catch {}
  };
  const getFlag = (id, flag) => !!payFlags[`${id}_${flag}`];
  const [activeView, setActiveView] = useState("unpaid-purchase");
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [payModal, setPayModal] = useState(null); // { kind: "purchase"|"sale", doc }

  const custName = (id) => customers.find((c) => c.id === id)?.name || id;

  // ---------- รายการใบรับสินค้าทั้งหมด (รวมที่ชำระครบแล้ว) ----------
  const allPurchaseRows = useMemo(() => {
    return purchases
      .filter((po) => (po.status || "") !== "ยกเลิก")
      .map((po) => {
        const subtotal = po.items.reduce((s, it) => {
          const qty = Number(it.qty) || 0;
          const deduct = Number(it.deduct) || 0;
          const net = it.deductType === "pct" ? qty * (1 - deduct / 100) : (it.net != null ? Number(it.net) : qty - deduct);
          const discountPct = Number(it.discountPct) || 0;
          return s + net * (Number(it.price) || 0) * (1 - discountPct / 100);
        }, 0);
        const vat = subtotal * ((Number(po.vatRate) || 0) / 100);
        const total = subtotal + vat;
        const paid = (po.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
        const remaining = total - paid;
        const payStatus = po.writeOff ? "paid" : (remaining > 0.01 ? (paid > 0.01 ? "partial" : "unpaid") : "paid");
        return { kind: "purchase", id: po.id, date: po.date, customerId: po.customerId, total, paid, remaining, payStatus, doc: po };
      });
  }, [purchases]);

  // ---------- รายการใบขายทั้งหมด (รวมที่ชำระครบแล้ว) ----------
  const allSaleRows = useMemo(() => {
    return sales
      .map((inv) => {
        const subtotal = inv.items.reduce((s, it) => s + (it.net || 0) * (it.price || 0), 0);
        const ad = subtotal - (inv.discount || 0);
        const vat = ad * ((Number(inv.vatRate) || 0) / 100);
        const total = ad + vat;
        const paid = (inv.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
        const remaining = total - paid;
        const payStatus = inv.writeOff ? "paid" : (remaining > 0.01 ? (paid > 0.01 ? "partial" : "unpaid") : "paid");
        return { kind: "sale", id: inv.id, date: inv.date, customerId: inv.customerId, total, paid, remaining, payStatus, doc: inv };
      });
  }, [sales]);

  // ---------- รายการค่าใช้จ่ายทั้งหมด (รวมที่ชำระครบแล้ว) ----------
  const allExpenseRows = useMemo(() => {
    return (expenses || []).map((e) => {
      const items = (e.items && e.items.length > 0) ? e.items : [{ amount: e.amount, vatEnabled: e.vatEnabled, whtRate: e.whtRate }];
      let amount = 0, vat = 0, wht = 0;
      items.forEach((it) => {
        const itAmount = Number(it.amount) || 0;
        amount += itAmount;
        vat += it.vatEnabled ? itAmount * 0.07 : 0;
        wht += itAmount * ((Number(it.whtRate) || 0) / 100);
      });
      const total = amount + vat - wht;
      const paid = (e.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
      const remaining = total - paid;
      const payStatus = e.writeOff ? "paid" : (remaining > 0.01 ? (paid > 0.01 ? "partial" : "unpaid") : "paid");
      const vendorLabel = e.vendorName || "";
      return { kind: "expense", id: e.refNo || e.id, date: e.billDate || e.recordDate || e.date, customerId: null, vendorLabel, total, paid, remaining, payStatus, doc: e };
    });
  }, [expenses]);

  // คำนวณยอดวงเงิน (ต้องอยู่หลัง allPurchaseRows/allSaleRows/allExpenseRows)
  const creditBalance = useMemo(() => {
    if (!creditLimit) return null;
    const getWithdrawn = (r) => !!payFlags[`${r.id}_withdrawn`];
    const totalBuy = allPurchaseRows.filter(r => getWithdrawn(r)).reduce((s, r) => s + r.total, 0);
    const totalExp = allExpenseRows.filter(r => getWithdrawn(r)).reduce((s, r) => s + r.total, 0);
    const totalSale = allSaleRows.filter(r => getWithdrawn(r)).reduce((s, r) => s + r.total, 0);
    const netOut = totalBuy + totalExp - totalSale;
    const balance = creditLimit - netOut;
    const pendingBuy = allPurchaseRows.filter(r => r.payStatus === "paid" && !getWithdrawn(r)).reduce((s, r) => s + r.total, 0);
    const pendingExp = allExpenseRows.filter(r => r.payStatus === "paid" && !getWithdrawn(r)).reduce((s, r) => s + r.total, 0);
    const pendingSale = allSaleRows.filter(r => r.payStatus === "paid" && !getWithdrawn(r)).reduce((s, r) => s + r.total, 0);
    const pendingNet = pendingBuy + pendingExp - pendingSale;
    return { limit: creditLimit, netOut, balance, pendingNet, totalBuy, totalExp, totalSale };
  }, [creditLimit, payFlags, allPurchaseRows, allExpenseRows, allSaleRows]);

  // คำนวณยอดรายวัน — เฉพาะรายการที่ชำระครบแล้ว และมี payment ผ่านบัญชีที่เลือก
  const creditDaySummary = useMemo(() => {
    const accs = new Set(creditAccounts);
    const hasAccPayment = (doc, field) =>
      accs.size === 0 || (doc.payments||[]).some(p => accs.has(p[field]));

    const dayCost = allPurchaseRows
      .filter(r => r.date === creditDate && r.payStatus === "paid" && hasAccPayment(r.doc, "fromStoreBankId"))
      .reduce((s,r)=>s+r.total,0);
    const dayExp  = allExpenseRows
      .filter(r => r.date === creditDate && r.payStatus === "paid" && hasAccPayment(r.doc, "fromStoreBankId"))
      .reduce((s,r)=>s+r.total,0);
    const dayRev  = allSaleRows
      .filter(r => r.date === creditDate && r.payStatus === "paid" && hasAccPayment(r.doc, "toStoreBankId"))
      .reduce((s,r)=>s+r.total,0);
    const dayNet  = dayCost + dayExp - dayRev;

    const pendingBefore =
      allPurchaseRows.filter(r => r.date < creditDate && r.payStatus==="paid" && !payFlags[`${r.id}_withdrawn`] && hasAccPayment(r.doc,"fromStoreBankId")).reduce((s,r)=>s+r.total,0)
      + allExpenseRows.filter(r => r.date < creditDate && r.payStatus==="paid" && !payFlags[`${r.id}_withdrawn`] && hasAccPayment(r.doc,"fromStoreBankId")).reduce((s,r)=>s+r.total,0)
      - allSaleRows.filter(r => r.date < creditDate && r.payStatus==="paid" && !payFlags[`${r.id}_withdrawn`] && hasAccPayment(r.doc,"toStoreBankId")).reduce((s,r)=>s+r.total,0);

    const manual = Number(creditManual) || 0;
    return { dayCost, dayExp, dayRev, dayNet, pendingBefore, manual, total: dayNet + pendingBefore + manual };
  }, [creditDate, allPurchaseRows, allExpenseRows, allSaleRows, payFlags, creditManual, creditAccounts]);

  const unpaidPurchases = allPurchaseRows.filter((r) => r.payStatus !== "paid");
  const unpaidSales = allSaleRows.filter((r) => r.payStatus !== "paid");
  const unpaidExpenses = allExpenseRows.filter((r) => r.payStatus !== "paid");

  const rowSearchLabel = (r) => r.kind === "expense" ? (r.vendorLabel || "") : custName(r.customerId);

  const combined = useMemo(() => {
    let list = [...allPurchaseRows, ...allSaleRows];
    if (activeView === "purchase") list = allPurchaseRows.filter((r) => r.payStatus === "paid");
    if (activeView === "sale") list = allSaleRows.filter((r) => r.payStatus === "paid");
    if (activeView === "expense") list = allExpenseRows.filter((r) => r.payStatus === "paid");
    if (activeView === "unpaid-purchase") list = allPurchaseRows.filter((r) => r.payStatus !== "paid");
    if (activeView === "unpaid-sale") list = allSaleRows.filter((r) => r.payStatus !== "paid");
    if (activeView === "unpaid-expense") list = allExpenseRows.filter((r) => r.payStatus !== "paid");
    return list
      .filter((r) => r.id.includes(search) || rowSearchLabel(r).includes(search))
      .filter((r) => (!dateFrom || (r.date || "") >= dateFrom) && (!dateTo || (r.date || "") <= dateTo))
      .sort((a, b) => {
        if (a.date !== b.date) return a.date < b.date ? 1 : -1;
        return a.id < b.id ? 1 : a.id > b.id ? -1 : 0;
      });
  }, [allPurchaseRows, allSaleRows, allExpenseRows, activeView, search, dateFrom, dateTo, customers]);

  const totalPayable = unpaidPurchases.reduce((s, r) => s + r.remaining, 0);
  const totalReceivable = unpaidSales.reduce((s, r) => s + r.remaining, 0);
  const totalPayableExpense = unpaidExpenses.reduce((s, r) => s + r.remaining, 0);

  // ---------- ฟอร์มบันทึกการจ่าย/รับเงิน (รองรับแบ่งจ่ายหลายงวดในครั้งเดียว) ----------
  const blankPaymentRow = (row, isFirst) => ({
    id: (row.kind === "sale" ? "SP" : "PM") + Date.now().toString().slice(-6) + Math.floor(Math.random() * 1000),
    date: new Date().toISOString().slice(0, 10),
    amount: isFirst ? Math.round(row.remaining * 100) / 100 : 0,
    method: PAYMENT_METHODS[0],
    fromStoreBankId: row.kind === "purchase" || row.kind === "expense" ? (storeBankAccounts[0]?.id || "") : undefined,
    toStoreBankId: row.kind === "sale" ? "" : undefined,
  });
  const [payRows, setPayRows] = useState(null);

  const openPay = (row) => {
    setPayModal(row);
    setPayRows([blankPaymentRow(row, true)]);
    setWriteOffChecked(false);
  };

  const addPayRow = () => setPayRows([...payRows, blankPaymentRow(payModal, false)]);
  const updatePayRow = (idx, field, value) => {
    const rows = [...payRows];
    rows[idx] = { ...rows[idx], [field]: value };
    setPayRows(rows);
  };
  const removePayRow = (idx) => setPayRows(payRows.filter((_, i) => i !== idx));

  // ยอดมัดจำคงเหลือของลูกค้า (เฉพาะกรณีจ่ายใบรับสินค้า — เผื่อหักมัดจำ)
  const depositBalanceForCustomer = (customerId, excludePoId) => {
    const opening = Number(customers.find((c) => c.id === customerId)?.depositOpening) || 0;
    const totalGiven = opening + (deposits || []).filter((d) => d.customerId === customerId).reduce((s, d) => s + (Number(d.amount) || 0), 0);
    const totalUsedOtherPOs = purchases
      .filter((po) => po.customerId === customerId && po.id !== excludePoId)
      .reduce((s, po) => s + (po.payments || []).filter((p) => p.fromStoreBankId === "DEPOSIT").reduce((s2, p) => s2 + (Number(p.amount) || 0), 0), 0);
    return totalGiven - totalUsedOtherPOs;
  };

  const rowsTotal = (payRows || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);

  const [writeOffChecked, setWriteOffChecked] = useState(false);

  const savePayment = () => {
    if (!payModal || !payRows || payRows.length === 0) return;
    const cleaned = payRows.filter((p) => Number(p.amount) > 0).map((p) => ({ ...p, amount: Number(p.amount) }));
    if (cleaned.length === 0) return;
    const realId = payModal.doc?.id ?? payModal.id;
    const ts = new Date().toISOString();
    if (payModal.kind === "purchase") {
      setPurchases(purchases.map((po) => po.id === realId ? { ...po, payments: [...(po.payments || []), ...cleaned], writeOff: writeOffChecked, updated_at: ts } : po));
    } else if (payModal.kind === "expense") {
      setExpenses(expenses.map((e) => e.id === realId ? { ...e, payments: [...(e.payments || []), ...cleaned], writeOff: writeOffChecked, updated_at: ts } : e));
    } else {
      setSales(sales.map((inv) => inv.id === realId ? { ...inv, payments: [...(inv.payments || []), ...cleaned], writeOff: writeOffChecked, updated_at: ts } : inv));
    }
    setPayModal(null);
    setPayRows(null);
    setWriteOffChecked(false);
  };

  const availableDeposit = payModal && payModal.kind === "purchase" ? depositBalanceForCustomer(payModal.customerId, payModal.id) : 0;

  // ---------- ดูประวัติ/แก้ไข/ลบ รายการที่บันทึกไปแล้ว ----------
  const [historyModal, setHistoryModal] = useState(null); // row ของใบที่กำลังดูประวัติ
  const [editingPaymentIdx, setEditingPaymentIdx] = useState(null); // index ของงวดที่กำลังแก้ไขอยู่ใน historyModal
  const [editPaymentForm, setEditPaymentForm] = useState(null);

  const openHistory = (row) => setHistoryModal(row);

  const startEditPayment = (idx, p) => {
    setEditingPaymentIdx(idx);
    setEditPaymentForm({ ...p });
  };

  const saveEditPayment = () => {
    if (!historyModal || editingPaymentIdx === null || !editPaymentForm) return;
    const cleaned = { ...editPaymentForm, amount: Number(editPaymentForm.amount) || 0 };
    const realId = historyModal.doc?.id ?? historyModal.id;
    if (historyModal.kind === "purchase") {
      setPurchases(purchases.map((po) => {
        if (po.id !== realId) return po;
        const newPayments = [...(po.payments || [])];
        newPayments[editingPaymentIdx] = cleaned;
        return { ...po, payments: newPayments };
      }));
    } else if (historyModal.kind === "expense") {
      setExpenses(expenses.map((e) => {
        if (e.id !== realId) return e;
        const newPayments = [...(e.payments || [])];
        newPayments[editingPaymentIdx] = cleaned;
        return { ...e, payments: newPayments };
      }));
    } else {
      setSales(sales.map((inv) => {
        if (inv.id !== realId) return inv;
        const newPayments = [...(inv.payments || [])];
        newPayments[editingPaymentIdx] = cleaned;
        return { ...inv, payments: newPayments };
      }));
    }
    // อัปเดต historyModal ให้ตรงกับข้อมูลใหม่ทันที (เพื่อให้ยอดสรุปในหน้าต่างอัปเดตด้วย)
    setHistoryModal({ ...historyModal, doc: { ...historyModal.doc, payments: (historyModal.doc.payments || []).map((p, i) => i === editingPaymentIdx ? cleaned : p) } });
    setEditingPaymentIdx(null);
    setEditPaymentForm(null);
  };

  const deleteHistoryPayment = (idx) => {
    if (!historyModal) return;
    const realId = historyModal.doc?.id ?? historyModal.id;
    const kind = historyModal.kind;
    const newPayments = (historyModal.doc.payments || []).filter((_, i) => i !== idx);
    const ts = new Date().toISOString();
    let updatedItem;
    if (kind === "purchase") {
      updatedItem = { ...historyModal.doc, payments: newPayments, writeOff: false, updated_at: ts };
      setPurchases(prev => prev.map((po) => po.id === realId ? updatedItem : po));
      saveToSupabase("purchases", [updatedItem]);
    } else if (kind === "expense") {
      updatedItem = { ...historyModal.doc, payments: newPayments, writeOff: false, updated_at: ts };
      setExpenses(prev => prev.map((e) => e.id === realId ? updatedItem : e));
      saveToSupabase("expenses", [updatedItem]);
    } else {
      updatedItem = { ...historyModal.doc, payments: newPayments, writeOff: false, updated_at: ts };
      setSales(prev => prev.map((inv) => inv.id === realId ? updatedItem : inv));
      saveToSupabase("sales", [updatedItem]);
    }
    // อัปเดต historyModal ให้แสดงยอดใหม่
    setHistoryModal({ ...historyModal, doc: updatedItem });
  };

  // คำนวณยอดล่าสุดของใบที่กำลังดูประวัติ (อ้างจาก historyModal.doc ที่อัปเดตสดๆ)
  const historyTotals = useMemo(() => {
    if (!historyModal) return null;
    const doc = historyModal.doc;
    let total;
    if (historyModal.kind === "expense") {
      const items = (doc.items && doc.items.length > 0) ? doc.items : [{ amount: doc.amount, vatEnabled: doc.vatEnabled, whtRate: doc.whtRate }];
      total = items.reduce((s, it) => {
        const itAmount = Number(it.amount) || 0;
        const vat = it.vatEnabled ? itAmount * 0.07 : 0;
        const wht = itAmount * ((Number(it.whtRate) || 0) / 100);
        return s + itAmount + vat - wht;
      }, 0);
    } else if (historyModal.kind === "purchase") {
      const subtotal = doc.items.reduce((s, it) => s + (it.deductType === "pct" ? (Number(it.qty)||0)*(1-(Number(it.deduct)||0)/100) : (it.net != null ? Number(it.net) : (Number(it.qty)||0)-(Number(it.deduct)||0))) * (Number(it.price)||0) * (1-(Number(it.discountPct)||0)/100), 0);
      total = subtotal + subtotal * ((Number(doc.vatRate) || 0) / 100);
    } else {
      const subtotal = doc.items.reduce((s, it) => s + (it.net || 0) * (it.price || 0), 0);
      const ad = subtotal - (doc.discount || 0);
      total = ad + ad * ((Number(doc.vatRate) || 0) / 100);
    }
    const paid = (doc.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
    return { total, paid, remaining: total - paid };
  }, [historyModal]);

  return (
    <div>
      <Header title="รับชำระ / จ่ายชำระ" subtitle="รวมรายการใบรับสินค้าและใบขายที่ยังค้างชำระ — บันทึกการจ่าย/รับเงินจริงได้ที่นี่" />

      {/* วงเงินหมุนเวียน */}
      {/* ตารางสรุปรายวัน */}
      <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "hidden", marginBottom: 16 }}>
        <div style={{ background: "#4a1e1e", color: "#fff", padding: "10px 16px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <span style={{ fontWeight: 700, fontSize: 14 }}>สรุปยอดใช้เงิน</span>
            <input type="date" value={creditDate} onChange={(e) => setCreditDate(e.target.value)}
              style={{ background: "rgba(255,255,255,0.15)", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 6, color: "#fff", padding: "3px 8px", fontSize: 13 }} />
          </div>
          <button onClick={() => printAsPDF("credit-day-summary-print", `สรุปยอดใช้เงิน ${creditDate}`)}
            style={{ background: "rgba(255,255,255,0.15)", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 6, color: "#fff", padding: "4px 12px", fontSize: 12, cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}>
            <Printer size={13} /> พิมพ์
          </button>
        </div>
        <div id="credit-day-summary-print">
          <div style={{ padding: "8px 14px", fontWeight: 700, fontSize: 14, borderBottom: "1px solid #e5e7eb" }}>
            สรุปยอดใช้เงิน — {creditDate}
          </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 0 }}>
          {/* ซ้าย: ยอดใช้เงินต่อวัน */}
          <div>
            <div style={{ background: "#6b1f1f", color: "#fff", padding: "6px 14px", fontSize: 12, fontWeight: 700 }}>ยอดใช้เงินต่อวัน / ยอดรับต่อวัน</div>
            {[
              { label: "ค่าสินค้า", value: creditDaySummary.dayCost, color: "#faece7" },
              { label: "ค่าใช้จ่าย", value: creditDaySummary.dayExp, color: "#fff" },
              { label: "หัก รายได้จากสินค้า", value: -creditDaySummary.dayRev, color: "#faece7" },
              { label: "รวมยอดใช้วันนี้", value: creditDaySummary.dayNet, color: "#e8d4d4", bold: true },
            ].map((row, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "7px 14px", background: row.color, borderBottom: "1px solid #f3f0f0" }}>
                <span style={{ fontSize: 13, fontWeight: row.bold ? 700 : 400 }}>{row.label}</span>
                <span style={{ fontSize: 13, fontWeight: row.bold ? 700 : 600, color: row.value < 0 ? "#0f6e56" : row.value > 0 ? "#993c1d" : "#374151" }}>
                  {row.value < 0 ? `(${fmt(Math.abs(row.value))})` : fmt(row.value)}
                </span>
              </div>
            ))}
          </div>
          {/* ขวา: ยอดที่ต้องเบิกคืน */}
          <div style={{ borderLeft: "1px solid #e5e7eb" }}>
            <div style={{ background: "#1a3a5c", color: "#fff", padding: "6px 14px", fontSize: 12, fontWeight: 700 }}>ยอดใช้ที่ต้องเบิกคืน</div>
            {[
              { label: "ยอดใช้วันนี้", value: creditDaySummary.dayNet, color: "#e6f1fb" },
              { label: "ยอดค้างเบิก", value: creditDaySummary.pendingBefore, color: "#fff" },
              { label: "ยอดตกหล่น", value: null, color: "#e6f1fb", input: true },
              { label: "ยอดรวมที่ต้องเบิก", value: creditDaySummary.total, color: "#d0e4f7", bold: true },
            ].map((row, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 14px", background: row.color, borderBottom: "1px solid #f0f4f8" }}>
                <span style={{ fontSize: 13, fontWeight: row.bold ? 700 : 400 }}>{row.label}</span>
                {row.input ? (
                  <input type="number" value={creditManual} onChange={(e) => setCreditManual(e.target.value)}
                    style={{ width: 100, textAlign: "right", border: "1px solid #d1d5db", borderRadius: 6, padding: "2px 8px", fontSize: 13 }} />
                ) : (
                  <span style={{ fontSize: 13, fontWeight: row.bold ? 700 : 600, color: row.value < 0 ? "#0f6e56" : row.value > 0 ? "#993c1d" : "#374151" }}>
                    {row.value < 0 ? `(${fmt(Math.abs(row.value))})` : fmt(row.value)}
                  </span>
                )}
              </div>
            ))}
          </div>
        </div>
        {/* ธนาคารที่จะโอนคืนวงเงิน */}
        <div style={{ borderTop: "1px solid #e5e7eb", padding: "10px 14px", background: "#f9fafb" }}>
          <div style={{ fontWeight: 600, fontSize: 13, color: "#374151", marginBottom: 8 }}>โอนคืนวงเงินผ่าน</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 12px" }}>
            <div>
              <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 3 }}>ธนาคาร</div>
              <input style={inputStyle} value={returnBankName} onChange={(e) => setReturnBankName(e.target.value)} placeholder="เช่น กสิกรไทย" />
            </div>
            <div>
              <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 3 }}>เลขที่บัญชี</div>
              <input style={inputStyle} value={returnBankNo} onChange={(e) => setReturnBankNo(e.target.value)} placeholder="xxx-x-xxxxx-x" />
            </div>
            <div>
              <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 3 }}>ชื่อบัญชี</div>
              <input style={inputStyle} value={returnBankOwner} onChange={(e) => setReturnBankOwner(e.target.value)} placeholder="ชื่อเจ้าของบัญชี" />
            </div>
          </div>
          {(returnBankName || returnBankNo) && (
            <div style={{ marginTop: 8, fontSize: 13, color: "#185fa5", fontWeight: 600 }}>
              ยอดโอนคืน: ฿{fmt(creditDaySummary.total)}
            </div>
          )}
        </div>
        </div>{/* end credit-day-summary-print */}
      </div>



      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginBottom: 20 }}>
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "16px 18px" }}>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>ยอดที่ต้องจ่าย (ใบรับสินค้าค้างจ่าย)</div>
          <div style={{ fontWeight: 700, fontSize: 22, color: "#993c1d" }}>฿{fmt(totalPayable)}</div>
          <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 4 }}>{unpaidPurchases.length} ใบ</div>
        </div>
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "16px 18px" }}>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>ยอดที่ต้องรับ (ใบขายค้างรับ)</div>
          <div style={{ fontWeight: 700, fontSize: 22, color: "#185fa5" }}>฿{fmt(totalReceivable)}</div>
          <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 4 }}>{unpaidSales.length} ใบ</div>
        </div>
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "16px 18px" }}>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>ยอดที่ต้องจ่าย (ค่าใช้จ่ายค้างจ่าย)</div>
          <div style={{ fontWeight: 700, fontSize: 22, color: "#854f0b" }}>฿{fmt(totalPayableExpense)}</div>
          <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 4 }}>{unpaidExpenses.length} รายการ</div>
        </div>
      </div>

      <div style={{ display: "flex", gap: 8, marginBottom: 16, overflowX: "auto", overflowY: "hidden", paddingBottom: 4 }}>
        {[
          { key: "unpaid-purchase", label: `ค้างจ่าย (ใบรับสินค้า) (${unpaidPurchases.length})` },
          { key: "unpaid-sale", label: `ค้างรับ (ใบขาย) (${unpaidSales.length})` },
          { key: "unpaid-expense", label: `ค้างจ่าย (ค่าใช้จ่าย) (${unpaidExpenses.length})` },
          { key: "purchase", label: `จ่ายชำระแล้ว (${allPurchaseRows.filter(r=>r.payStatus==="paid").length})` },
          { key: "sale", label: `รับชำระแล้ว (${allSaleRows.filter(r=>r.payStatus==="paid").length})` },
          { key: "expense", label: `ค่าใช้จ่ายชำระแล้ว (${allExpenseRows.filter(r=>r.payStatus==="paid").length})` },
        ].map((opt) => (
          <button key={opt.key} onClick={() => setActiveView(opt.key)}
            style={{ flexShrink: 0, padding: "8px 16px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, border: "1px solid",
              borderColor: activeView === opt.key ? "#1d9e75" : "#d1d5db",
              background: activeView === opt.key ? "#e1f5ee" : "#fff",
              color: activeView === opt.key ? "#0f6e56" : "#6b7280" }}>
            {opt.label}
          </button>
        ))}
        {/* ปุ่มสรุปตั้งโอน */}
        {(() => {
          const transferList = [...unpaidPurchases, ...unpaidSales, ...unpaidExpenses].filter(r => getFlag(r.id, "transfer"));
          return (
            <button onClick={() => setShowTransferSheet(true)}
              style={{ flexShrink: 0, padding: "8px 16px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600,
                border: "1px solid #185fa5", background: "#185fa5", color: "#fff", display: "flex", alignItems: "center", gap: 6 }}>
              <FileText size={14} /> สรุปตั้งโอน ({transferList.length} รายการ)
            </button>
          );
        })()}
      </div>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาเลขที่ใบ หรือชื่อลูกค้า..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />

      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={thStyle}>ประเภท</th>
              <th style={thStyle}>เลขที่ใบ</th>
              <th style={thStyle}>วันที่</th>
              <th style={thStyle}>ลูกค้า / รายการ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ยอดรวม</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ชำระแล้ว</th>
              <th style={{ ...thStyle, textAlign: "right" }}>คงค้าง</th>
              {["unpaid-purchase","unpaid-sale","unpaid-expense"].includes(activeView) && (
                <th style={{ ...thStyle, textAlign: "center" }}>ตั้งโอน</th>
              )}
              {["purchase","sale","expense"].includes(activeView) && (
                <th style={{ ...thStyle, textAlign: "center" }}>เบิกแล้ว</th>
              )}
              <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {combined.map((r) => (
              <tr key={r.kind + r.id}>
                <td style={tdStyle}>
                  {r.kind === "purchase" ? (
                    <span style={{ background: "#faece7", color: "#993c1d", padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 600 }}>จ่าย (ใบรับสินค้า)</span>
                  ) : r.kind === "expense" ? (
                    <span style={{ background: "#faeeda", color: "#854f0b", padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 600 }}>จ่าย (ค่าใช้จ่าย)</span>
                  ) : (
                    <span style={{ background: "#e6f1fb", color: "#185fa5", padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 600 }}>รับ (ใบขาย)</span>
                  )}
                </td>
                <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{r.id}</td>
                <td style={tdStyle}>{r.date}</td>
                <td style={tdStyle}>{r.kind === "expense" ? r.vendorLabel : custName(r.customerId)}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(r.total)}</td>
                <td style={{ ...tdStyle, textAlign: "right", color: "#0f6e56" }}>฿{fmt(r.paid)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: r.payStatus === "paid" ? "#0f6e56" : r.kind === "sale" ? "#185fa5" : "#993c1d" }}>฿{fmt(r.payStatus === "paid" ? 0 : r.remaining)}</td>

                {["unpaid-purchase","unpaid-sale","unpaid-expense"].includes(activeView) && (
                  <td style={{ ...tdStyle, textAlign: "center" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                      <input type="checkbox" checked={getFlag(r.id, "transfer")} onChange={(e) => {
                        setFlag(r.id, "transfer", e.target.checked);
                        if (e.target.checked) {
                          const existing = transferDetails[r.id];
                          setTransferEntries(existing?.length > 0 ? existing.map(e=>({...e, amount: String(e.amount)})) : [{ bankId: "", amount: String(r.remaining) }]);
                          setTransferDetailModal({ row: r });
                        }
                      }} style={{ width: 16, height: 16, cursor: "pointer", accentColor: "#185fa5" }} />
                      {getFlag(r.id, "transfer") && transferDetails[r.id]?.bankId && (
                        <span style={{ fontSize: 10, color: "#185fa5", cursor: "pointer" }}
                          onClick={() => { setTransferBankId(transferDetails[r.id]?.bankId || ""); setTransferAmount(transferDetails[r.id]?.amount || r.remaining); setTransferDetailModal({ row: r }); }}>
                          {(storeBankAccounts.find(a=>a.id===transferDetails[r.id].bankId)?.bankName||"") + " ฿" + fmt(transferDetails[r.id].amount)}
                        </span>
                      )}
                    </div>
                  </td>
                )}
                {["purchase","sale","expense"].includes(activeView) && (
                  <td style={{ ...tdStyle, textAlign: "center" }}>
                    <input type="checkbox" checked={getFlag(r.id, "withdrawn")} onChange={(e) => {
                        if (!e.target.checked && getFlag(r.id, "withdrawn")) {
                          confirmAction(`ต้องการยกเลิก "เบิกแล้ว" ของรายการ ${r.id} ใช่หรือไม่?`, () => setFlag(r.id, "withdrawn", false));
                        } else {
                          setFlag(r.id, "withdrawn", e.target.checked);
                        }
                      }}
                      style={{ width: 16, height: 16, cursor: "pointer", accentColor: "#0f6e56" }} />
                  </td>
                )}
                <td style={{ ...tdStyle, textAlign: "right" }}>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    {r.payStatus !== "paid" && (
                      <button style={btnPrimary} onClick={() => openPay(r)}>
                        {r.kind === "sale" ? <><ArrowDownToLine size={14} /> บันทึกรับ</> : <><ArrowUpFromLine size={14} /> บันทึกจ่าย</>}
                      </button>
                    )}
                    {(r.paid > 0.01) && (
                      <button style={iconBtn} onClick={() => openHistory(r)}><History size={14} /> ประวัติ</button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
            {combined.length === 0 && <tr><td colSpan={9} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีรายการในตัวกรองนี้</td></tr>}
          </tbody>
        </table>
      </Card>

      {/* Modal ตั้งค่าวงเงินหมุนเวียน */}
      {showCreditSetting && (
        <Modal title="ตั้งค่าวงเงินหมุนเวียนร้าน" onClose={() => setShowCreditSetting(false)}>
          <Field label="วงเงินตั้งต้น (บาท)">
            <input type="number" style={{ ...inputStyle, textAlign: "right" }}
              value={companySettings?.creditLimit || ""}
              onChange={(e) => setCompanySettings(prev => ({ ...prev, creditLimit: Number(e.target.value) || 0 }))}
              placeholder="เช่น 500000" />
          </Field>
          <div style={{ marginTop: 16 }}>
            <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 8 }}>เลือกบัญชีที่นับในวงเงิน</div>
            {storeBankAccounts.map((acc) => (
              <label key={acc.id} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, fontSize: 13, cursor: "pointer" }}>
                <input type="checkbox"
                  checked={(companySettings?.creditAccounts || []).includes(acc.id)}
                  onChange={(e) => {
                    const cur = companySettings?.creditAccounts || [];
                    const next = e.target.checked ? [...cur, acc.id] : cur.filter(id => id !== acc.id);
                    setCompanySettings(prev => ({ ...prev, creditAccounts: next }));
                  }}
                  style={{ width: 16, height: 16 }} />
                {acc.bankName} — {acc.accountNo} ({acc.accountName || ""})
              </label>
            ))}
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
            <button style={btnPrimary} onClick={() => setShowCreditSetting(false)}><Check size={14} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {/* Modal ตั้งโอน — เลือกบัญชีและยอด */}
      {transferDetailModal && (
        <Modal title="ตั้งโอน — เลือกบัญชีและยอด" onClose={() => setTransferDetailModal(null)}>
          <div style={{ marginBottom: 12, fontSize: 13, color: "#6b7280" }}>
            เลขที่: <strong>{transferDetailModal.row.id}</strong> — {custName(transferDetailModal.row.customerId) || transferDetailModal.row.vendorLabel}
          </div>
          <div style={{ marginBottom: 12, fontSize: 13 }}>
            ยอดคงค้าง: <strong style={{ color: "#993c1d" }}>฿{fmt(transferDetailModal.row.remaining)}</strong>
          </div>
          {transferEntries.map((entry, idx) => (
            <div key={idx} style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: "0 10px", alignItems: "end", marginBottom: 10 }}>
              <Field label={`บัญชีที่ ${idx+1}`}>
                <select style={inputStyle} value={entry.bankId} onChange={(e) => {
                  const next = [...transferEntries]; next[idx] = {...next[idx], bankId: e.target.value}; setTransferEntries(next);
                }}>
                  <option value="">-- เลือกบัญชี --</option>
                  {(() => {
                      const cust = customers.find(c => c.id === (transferDetailModal?.row?.customerId || transferDetailModal?.row?.doc?.vendorId));
                      const bankAccs = cust?.bankAccounts || [];
                      return bankAccs.length > 0
                        ? bankAccs.map((b, bi) => <option key={bi} value={JSON.stringify({bankName:b.bankName,accountNo:b.accountNo})}>{b.bankName} — {b.accountNo}</option>)
                        : <option value="" disabled>ไม่มีบัญชีลูกค้า</option>;
                    })()}
                </select>
              </Field>
              <Field label="ยอด (บาท)">
                <input type="number" style={{ ...inputStyle, textAlign: "right" }} value={entry.amount}
                  onChange={(e) => { const next=[...transferEntries]; next[idx]={...next[idx],amount:e.target.value}; setTransferEntries(next); }}
                  placeholder="0" />
              </Field>
              <div style={{ paddingBottom: 4 }}>
                {transferEntries.length > 1 && (
                  <button style={{ ...btnDanger, padding: "6px 10px" }} onClick={() => setTransferEntries(transferEntries.filter((_,i)=>i!==idx))}>
                    <X size={14} />
                  </button>
                )}
              </div>
            </div>
          ))}
          <button style={{ ...btnSecondary, marginBottom: 12 }} onClick={() => setTransferEntries([...transferEntries, { bankId: "", amount: "" }])}>
            <Plus size={13} /> เพิ่มบัญชี
          </button>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
            <span style={{ fontSize: 13, color: "#6b7280" }}>
              รวมยอดโอน: <strong style={{ color: "#185fa5" }}>฿{fmt(transferEntries.reduce((s,e)=>s+(Number(e.amount)||0),0))}</strong>
            </span>
            <div style={{ display: "flex", gap: 8 }}>
              <button style={btnSecondary} onClick={() => setTransferDetailModal(null)}>ยกเลิก</button>
              <button style={btnPrimary} onClick={() => {
                saveTransferDetail(transferDetailModal.row.id, transferEntries.map(e=>({bankId:e.bankId,amount:Number(e.amount)||0})));
                setTransferDetailModal(null);
              }}><Check size={14} /> บันทึก</button>
            </div>
          </div>
        </Modal>
      )}

      {/* Modal สรุปตั้งโอน */}
      {showTransferSheet && (() => {
        const transferList = [...unpaidPurchases, ...unpaidSales, ...unpaidExpenses].filter(r => getFlag(r.id, "transfer"));
        const totalAmt = transferList.reduce((s, r) => s + r.remaining, 0);
        return (
          <Modal title="สรุปรายการตั้งโอน" onClose={() => setShowTransferSheet(false)} wide>
            <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
              {[
                { key: "purchase", label: `ใบรับสินค้า (${transferList.filter(r=>r.kind==="purchase").length})` },
                { key: "expense",  label: `ค่าใช้จ่าย (${transferList.filter(r=>r.kind==="expense").length})` },
                { key: "sale",     label: `ใบขาย (${transferList.filter(r=>r.kind==="sale").length})` },
              ].map(opt => (
                <button key={opt.key} onClick={() => setTransferTab(opt.key)}
                  style={{ padding: "8px 18px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, border: "1px solid",
                    borderColor: transferTab === opt.key ? "#185fa5" : "#d1d5db",
                    background: transferTab === opt.key ? "#185fa5" : "#fff",
                    color: transferTab === opt.key ? "#fff" : "#6b7280" }}>
                  {opt.label}
                </button>
              ))}
            </div>
            <div id="transfer-sheet-print" style={{ fontSize: 12 }}>
              <style>{`
                @media print {
                  #transfer-sheet-print table { font-size: 9px !important; }
                  #transfer-sheet-print td, #transfer-sheet-print th { padding: 2px 4px !important; }
                }
              `}</style>
              <div style={{ marginBottom: 16, fontWeight: 700, fontSize: 15 }}>
                {transferTab === "purchase" ? "ตั้งโอนใบรับสินค้า" : transferTab === "expense" ? "ตั้งโอนค่าใช้จ่าย" : "ตั้งโอนใบขาย"} — {new Date().toLocaleDateString("th-TH")}
              </div>
              {transferList.filter(r => r.kind === transferTab).length === 0 ? (
                <p style={{ color: "#9ca3af", textAlign: "center", padding: 24 }}>ยังไม่มีรายการที่ติ๊กตั้งโอน</p>
              ) : (
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr>
                      <th style={thStyle}>เลขที่ใบ</th>
                      <th style={thStyle}>วันที่</th>
                      <th style={thStyle}>ลูกค้า / รายการ</th>
                      {transferTab === "expense" && <th style={thStyle}>รายละเอียดค่าใช้จ่าย</th>}
                      <th style={thStyle}>บัญชีลูกค้า</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>ยอดตั้งโอน</th>
                    </tr>
                  </thead>
                  {(() => {
                    const filteredByTab = transferList.filter(r => r.kind === transferTab);
                    // แสดงแถวตามข้อมูลตั้งโอนที่กรอกไว้
                    const expandToPaymentRows = (r) => {
                      const vendorId = r.doc?.vendorId;
                      const expenseDetail = r.kind === "expense" ? (() => {
                        const exp = expenses.find(e => (e.refNo || e.id) === r.id || e.id === r.id);
                        if (!exp) return "-";
                        const items = exp.items && exp.items.length > 0 ? exp.items : [{ subCategory: exp.subCategory, description: exp.description }];
                        return items.map(it => [it.subCategory, it.description].filter(Boolean).join(" — ")).join(", ");
                      })() : null;
                      const details = transferDetails[r.id] || [];
                      const totalTransfer = details.reduce((s,e)=>s+(Number(e.amount)||0),0) || r.remaining;
                      if (details.length <= 1) {
                        const acc = details[0]?.bankId ? storeBankAccounts.find(a=>a.id===details[0].bankId) : null;
                        const bankInfo = acc ? `${acc.bankName} — ${acc.accountNo}` : "-";
                        return [(
                          <tr key={r.id}>
                            <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{r.id}</td>
                            <td style={tdStyle}>{r.date}</td>
                            <td style={tdStyle}>{r.kind === "expense" ? r.vendorLabel : custName(r.customerId)}</td>
                            {transferTab === "expense" && <td style={{ ...tdStyle, fontSize: 12, color: "#6b7280" }}>{expenseDetail || "-"}</td>}
                            <td style={{ ...tdStyle, fontSize: 12, color: "#185fa5" }}>{bankInfo}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: r.kind === "sale" ? "#185fa5" : "#993c1d" }}>฿{fmt(details[0]?.amount || r.remaining)}</td>
                          </tr>
                        )];
                      }
                      // หลายบัญชี
                      return details.map((d, di) => {
                          let bankLabel = "-";
                          try {
                            const parsed = JSON.parse(d.bankId);
                            bankLabel = `${parsed.bankName} — ${parsed.accountNo}`;
                          } catch { bankLabel = d.bankId || "-"; }
                          return (
                            <tr key={`${r.id}-d${di}`}>
                              <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{di === 0 ? r.id : ""}</td>
                              <td style={tdStyle}>{di === 0 ? r.date : ""}</td>
                              <td style={tdStyle}>{di === 0 ? (r.kind === "expense" ? r.vendorLabel : custName(r.customerId)) : ""}</td>
                              {transferTab === "expense" && <td style={{ ...tdStyle, fontSize: 12, color: "#6b7280" }}>{di === 0 ? (expenseDetail || "-") : ""}</td>}
                              <td style={{ ...tdStyle, fontSize: 12, color: "#185fa5" }}>{bankLabel}</td>
                              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#185fa5" }}>฿{fmt(Number(d.amount)||0)}</td>
                            </tr>
                          );
                        });
                    };
                    const tabTotal = filteredByTab.reduce((s, r) => {
                      const dets = transferDetails[r.id] || [];
                      const amt = dets.reduce((s2,d)=>s2+(Number(d.amount)||0),0) || r.remaining;
                      return s + amt;
                    }, 0);

                    // จัดกลุ่มตามลูกค้า
                    const grouped = [];
                    const custMap = {};
                    filteredByTab.forEach(r => {
                      const key = r.customerId || r.vendorLabel || r.id;
                      if (!custMap[key]) { custMap[key] = []; grouped.push({ key, rows: custMap[key] }); }
                      custMap[key].push(r);
                    });

                    return (
                      <tbody>
                        {grouped.flatMap(({ key, rows }) => {
                          const custTotal = rows.reduce((s, r) => {
                            const dets = transferDetails[r.id] || [];
                            return s + (dets.reduce((s2,d)=>s2+(Number(d.amount)||0),0) || r.remaining);
                          }, 0);
                          const custName2 = rows[0].kind === "expense" ? rows[0].vendorLabel : custName(rows[0].customerId);
                          return [
                            ...rows.flatMap(expandToPaymentRows),
                            <tr key={`sum-${key}`} style={{ background: "#e6f1fb", borderTop: "1px solid #bfdbfe" }}>
                              <td colSpan={transferTab === "expense" ? 4 : 3} style={{ ...tdStyle, fontWeight: 700, color: "#185fa5", paddingLeft: 24 }}>
                                {custName2} — รวม
                              </td>
                              <td style={{ ...tdStyle }}></td>
                              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#185fa5", fontSize: 14 }}>฿{fmt(custTotal)}</td>
                            </tr>
                          ];
                        })}
                        <tr style={{ borderTop: "2px solid #185fa5" }}>
                          <td colSpan={transferTab === "expense" ? 5 : 4} style={{ ...tdStyle, fontWeight: 700, color: "#185fa5" }}>รวมยอดทั้งหมด</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 16, color: "#185fa5" }}>฿{fmt(tabTotal)}</td>
                        </tr>
                      </tbody>
                    );
                  })()}
                </table>
              )}
            </div>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
              <button style={btnSecondary} onClick={() => printAsPDF("transfer-sheet-print", "สรุปตั้งโอน")}><Printer size={14} /> พิมพ์ / PDF</button>
              <button style={btnSecondary} onClick={() => setShowTransferSheet(false)}>ปิด</button>
            </div>
          </Modal>
        );
      })()}

      {payModal && payRows && (
        <Modal
          title={payModal.kind === "sale" ? `บันทึกรับเงิน · ${payModal.id}` : `บันทึกจ่ายเงิน · ${payModal.id}`}
          onClose={() => { setPayModal(null); setPayRows(null); }}
          wide
        >
          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontSize: 13 }}>
            <Row label={payModal.kind === "expense" ? "รายการ" : "ลูกค้า"} value={payModal.kind === "expense" ? payModal.vendorLabel : custName(payModal.customerId)} />
            <Row label="ยอดรวมใบนี้" value={`฿${fmt(payModal.total)}`} />
            <Row label="ชำระแล้ว" value={`฿${fmt(payModal.paid)}`} />
            <Row label="คงค้าง" value={`฿${fmt(payModal.remaining)}`} bold color={payModal.kind === "sale" ? "#185fa5" : "#993c1d"} />
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontWeight: 600, fontSize: 14 }}>
              {payModal.kind === "sale" ? "รายการรับเงิน" : "รายการจ่ายเงิน"} (แบ่งจ่าย/รับได้หลายงวด)
            </div>
            <button style={btnSecondary} onClick={addPayRow}><Plus size={14} /> เพิ่มรายการ</button>
          </div>

          {payRows.map((p, idx) => (
            <div key={p.id} style={{ border: "1px solid #e5e7eb", borderRadius: 10, padding: 12, marginBottom: 10, background: "#f9fafb" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontWeight: 600, fontSize: 12, color: "#6b7280" }}>งวดที่ {idx + 1}</div>
                {payRows.length > 1 && (
                  <button style={btnDanger} onClick={() => removePayRow(idx)}><Trash2 size={14} /> ลบ</button>
                )}
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: "0 14px" }}>
                <Field label="วันที่">
                  <input type="date" style={inputStyle} value={p.date} onChange={(e) => updatePayRow(idx, "date", e.target.value)} />
                </Field>
                <Field label="จำนวนเงิน">
                  <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                    <input type="number" style={{ ...inputStyle, textAlign: "right" }} value={p.amount} onChange={(e) => updatePayRow(idx, "amount", e.target.value)} />
                    <button type="button" title="ปัดขึ้นเป็นจำนวนเต็มบาท" style={roundBtn} onClick={() => updatePayRow(idx, "amount", roundUpAmount(p.amount))}>▲</button>
                    <button type="button" title="ปัดลงเป็นจำนวนเต็มบาท" style={roundBtn} onClick={() => updatePayRow(idx, "amount", roundDownAmount(p.amount))}>▼</button>
                  </div>
                </Field>
              </div>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
                {payModal.kind === "purchase" ? (
                  <Field label="จ่ายจากบัญชี/วิธีจ่าย">
                    <select style={inputStyle} value={p.fromStoreBankId} onChange={(e) => updatePayRow(idx, "fromStoreBankId", e.target.value)}>
                      <option value="">-- เลือกบัญชี/วิธีจ่าย --</option>
                      <option value="DEPOSIT">หักเงินมัดจำ</option>
                      {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
                    </select>
                  </Field>
                ) : payModal.kind === "expense" ? (
                  <Field label="จ่ายจากบัญชี">
                    <select style={inputStyle} value={p.fromStoreBankId} onChange={(e) => updatePayRow(idx, "fromStoreBankId", e.target.value)}>
                      <option value="">-- เลือกบัญชี --</option>
                      {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
                    </select>
                  </Field>
                ) : (
                  <Field label="รับเข้าบัญชี">
                    <select style={inputStyle} value={p.toStoreBankId || ""} onChange={(e) => updatePayRow(idx, "toStoreBankId", e.target.value)}>
                      <option value="">เงินสด / ไม่ระบุบัญชี</option>
                      <option value="PREPAYMENT">หักจากรับล่วงหน้า</option>
                      {(storeBankAccounts || []).map((b) => <option key={b.id} value={b.id}>{b.bankName} · {b.accountNo}</option>)}
                    </select>
                  </Field>
                )}
                <Field label="วิธีชำระ">
                  <select style={inputStyle} value={p.method} onChange={(e) => updatePayRow(idx, "method", e.target.value)}>
                    {PAYMENT_METHODS.map((m) => <option key={m} value={m}>{m}</option>)}
                  </select>
                </Field>
              </div>

              {payModal.kind === "purchase" && p.fromStoreBankId === "DEPOSIT" && (
                <p style={{ fontSize: 12, color: (Number(p.amount) || 0) > availableDeposit ? "#a32d2d" : "#6b9c8d", margin: "4px 0 0" }}>
                  ลูกค้ามีเงินมัดจำคงเหลือ ฿{fmt(availableDeposit)}
                  {(Number(p.amount) || 0) > availableDeposit && " — เกินยอดมัดจำคงเหลือ"}
                </p>
              )}
            </div>
          ))}

          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "10px 14px", marginTop: 4, fontSize: 14 }}>
            <Row label="รวมยอดที่จะบันทึกครั้งนี้" value={`฿${fmt(rowsTotal)}`} bold />
            <Row label="คงค้างหลังบันทึก" value={`฿${fmt(payModal.remaining - rowsTotal)}`} color={(payModal.remaining - rowsTotal) > 0.01 ? "#993c1d" : "#0f6e56"} />
          </div>

          {Math.abs(payModal.remaining - rowsTotal) > 0.01 && (
            <label style={{ display: "flex", alignItems: "flex-start", gap: 8, marginTop: 10, padding: "10px 14px", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, cursor: "pointer", fontSize: 13 }}>
              <input type="checkbox" checked={writeOffChecked} onChange={(e) => setWriteOffChecked(e.target.checked)} style={{ marginTop: 2 }} />
              <span>
                <strong>ปัดเศษให้ครบ — ถือว่าใบนี้ "ชำระครบแล้ว"</strong>
                <br />
                <span style={{ color: "#854f0b" }}>ไม่นับยอดคงเหลือ ฿{fmt(payModal.remaining - rowsTotal)} เป็นยอดค้างอีกต่อไป (เหมาะกับเศษสตางค์เล็กน้อยจากการชั่งน้ำหนัก)</span>
              </span>
            </label>
          )}

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => { setPayModal(null); setPayRows(null); setWriteOffChecked(false); }}>ยกเลิก</button>
            <button style={btnPrimary} onClick={savePayment}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {historyModal && historyTotals && (
        <Modal
          title={`ประวัติการชำระเงิน · ${historyModal.id}`}
          onClose={() => { setHistoryModal(null); setEditingPaymentIdx(null); setEditPaymentForm(null); }}
          wide
        >
          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "10px 14px", marginBottom: 14, fontSize: 13 }}>
            <Row label={historyModal.kind === "expense" ? "รายการ" : "ลูกค้า"} value={historyModal.kind === "expense" ? historyModal.vendorLabel : custName(historyModal.customerId)} />
            <Row label="ยอดรวมใบนี้" value={`฿${fmt(historyTotals.total)}`} />
            <Row label="ชำระแล้ว" value={`฿${fmt(historyTotals.paid)}`} />
            <Row label="คงค้าง" value={`฿${fmt(historyTotals.remaining)}`} bold color={historyTotals.remaining > 0.01 ? "#993c1d" : "#0f6e56"} />
          </div>

          {historyModal.doc.writeOff && historyTotals.remaining > 0.01 && (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 14, padding: "10px 14px", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, fontSize: 13 }}>
              <span style={{ color: "#854f0b" }}>
                <strong>ปัดเศษไว้แล้ว</strong> — ใบนี้ถือว่าชำระครบ ไม่นับยอด ฿{fmt(historyTotals.remaining)} เป็นยอดค้าง
              </span>
              <button style={btnSecondary} onClick={() => {
                const realId = historyModal.doc?.id ?? historyModal.id;
                const ts = new Date().toISOString();
                const updatedDoc = { ...historyModal.doc, writeOff: false, updated_at: ts };
                if (historyModal.kind === "purchase") {
                  setPurchases(prev => prev.map((po) => po.id === realId ? updatedDoc : po));
                  saveToSupabase("purchases", [updatedDoc]);
                } else if (historyModal.kind === "expense") {
                  setExpenses(prev => prev.map((e) => e.id === realId ? updatedDoc : e));
                  saveToSupabase("expenses", [updatedDoc]);
                } else {
                  setSales(prev => prev.map((inv) => inv.id === realId ? updatedDoc : inv));
                  saveToSupabase("sales", [updatedDoc]);
                }
                setHistoryModal({ ...historyModal, doc: updatedDoc });
              }}>ยกเลิกการปัดเศษ</button>
            </div>
          )}

          <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 10 }}>
            {historyModal.kind === "sale" ? "ประวัติการรับเงิน" : "ประวัติการจ่ายเงิน"} — แก้ไขหรือลบงวดที่บันทึกผิดได้
          </div>

          {(historyModal.doc.payments || []).length === 0 && (
            <p style={{ color: "#9ca3af", fontSize: 13 }}>ยังไม่มีประวัติการชำระเงิน</p>
          )}

          {(historyModal.doc.payments || []).map((p, idx) => (
            <div key={p.id || idx} style={{ border: "1px solid #e5e7eb", borderRadius: 10, padding: 12, marginBottom: 10, background: "#f9fafb" }}>
              {editingPaymentIdx === idx && editPaymentForm ? (
                <>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: "0 14px" }}>
                    <Field label="วันที่">
                      <input type="date" style={inputStyle} value={editPaymentForm.date} onChange={(e) => setEditPaymentForm({ ...editPaymentForm, date: e.target.value })} />
                    </Field>
                    <Field label="จำนวนเงิน">
                      <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                        <input type="number" style={{ ...inputStyle, textAlign: "right" }} value={editPaymentForm.amount} onChange={(e) => setEditPaymentForm({ ...editPaymentForm, amount: e.target.value })} />
                        <button type="button" title="ปัดขึ้นเป็นจำนวนเต็มบาท" style={roundBtn} onClick={() => setEditPaymentForm({ ...editPaymentForm, amount: roundUpAmount(editPaymentForm.amount) })}>▲</button>
                        <button type="button" title="ปัดลงเป็นจำนวนเต็มบาท" style={roundBtn} onClick={() => setEditPaymentForm({ ...editPaymentForm, amount: roundDownAmount(editPaymentForm.amount) })}>▼</button>
                      </div>
                    </Field>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
                    {historyModal.kind === "purchase" ? (
                      <Field label="จ่ายจากบัญชี/วิธีจ่าย">
                        <select style={inputStyle} value={editPaymentForm.fromStoreBankId || ""} onChange={(e) => setEditPaymentForm({ ...editPaymentForm, fromStoreBankId: e.target.value })}>
                          <option value="">-- เลือกบัญชี/วิธีจ่าย --</option>
                          <option value="DEPOSIT">หักเงินมัดจำ</option>
                          {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
                        </select>
                      </Field>
                    ) : historyModal.kind === "expense" ? (
                      <Field label="จ่ายจากบัญชี">
                        <select style={inputStyle} value={editPaymentForm.fromStoreBankId || ""} onChange={(e) => setEditPaymentForm({ ...editPaymentForm, fromStoreBankId: e.target.value })}>
                          <option value="">-- เลือกบัญชี --</option>
                          {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
                        </select>
                      </Field>
                    ) : (
                      <Field label="รับเข้าบัญชี">
                        <select style={inputStyle} value={editPaymentForm.toStoreBankId || ""} onChange={(e) => setEditPaymentForm({ ...editPaymentForm, toStoreBankId: e.target.value })}>
                          <option value="">เงินสด / ไม่ระบุบัญชี</option>
                          <option value="PREPAYMENT">หักจากรับล่วงหน้า</option>
                          {(storeBankAccounts || []).map((b) => <option key={b.id} value={b.id}>{b.bankName} · {b.accountNo}</option>)}
                        </select>
                      </Field>
                    )}
                    <Field label="วิธีชำระ">
                      <select style={inputStyle} value={editPaymentForm.method} onChange={(e) => setEditPaymentForm({ ...editPaymentForm, method: e.target.value })}>
                        {PAYMENT_METHODS.map((m) => <option key={m} value={m}>{m}</option>)}
                      </select>
                    </Field>
                  </div>
                  <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 8 }}>
                    <button style={btnSecondary} onClick={() => { setEditingPaymentIdx(null); setEditPaymentForm(null); }}>ยกเลิก</button>
                    <button style={btnPrimary} onClick={saveEditPayment}><Save size={14} /> บันทึกการแก้ไข</button>
                  </div>
                </>
              ) : (
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontSize: 13 }}>
                    <div><strong>฿{fmt(p.amount)}</strong> — {p.date}</div>
                    <div style={{ color: "#6b7280", fontSize: 12, marginTop: 2 }}>
                      {p.method || ""}
                      {(historyModal.kind === "purchase" || historyModal.kind === "expense") && p.fromStoreBankId && (
                        <> · {p.fromStoreBankId === "CASH" ? "เงินสดหน้าร้าน" : p.fromStoreBankId === "DEPOSIT" ? "หักเงินมัดจำ" : (storeBankAccounts.find((b) => b.id === p.fromStoreBankId)?.bankName || "")}</>
                      )}
                      {historyModal.kind === "sale" && p.toStoreBankId && (
                        <> · {storeBankAccounts.find((b) => b.id === p.toStoreBankId)?.bankName || ""}</>
                      )}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button style={iconBtn} onClick={() => startEditPayment(idx, p)}><Edit2 size={14} /> แก้ไข</button>
                    <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบรายการชำระเงิน ฿${fmt(p.amount)} วันที่ ${p.date} ใช่หรือไม่?`, () => deleteHistoryPayment(idx))}><Trash2 size={14} /> ลบ</button>
                  </div>
                </div>
              )}
            </div>
          ))}

          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => { setHistoryModal(null); setEditingPaymentIdx(null); setEditPaymentForm(null); }}>ปิด</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ===================================================================
// INVENTORY TAB
// ===================================================================
function InventoryTab({ products, inventory, storeBankAccounts }) {
  const [expanded, setExpanded] = useState(null);
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [showBasket, setShowBasket] = useState(false);
  const today = new Date().toISOString().slice(0, 10);

  const totalOpeningStockQty   = products.reduce((s, p) => s + (Number(p.openingQty) || 0), 0);
  const totalOpeningStockValue = products.reduce((s, p) => s + (Number(p.openingQty) || 0) * (Number(p.openingCost) || 0), 0);
  const totalOpeningBankBalance = (storeBankAccounts || []).reduce((s, a) => s + (Number(a.openingBalance) || 0), 0);

  const toggleSelect = (productId) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(productId)) next.delete(productId);
      else next.add(productId);
      return next;
    });
  };
  const clearBasket = () => { setSelectedIds(new Set()); setShowBasket(false); };

  const basketItems = inventory.summary.filter((s) => selectedIds.has(s.productId));
  const basketTotalQty = basketItems.reduce((sum, s) => sum + (Number(s.qty) || 0), 0);
  const basketTotalValue = basketItems.reduce((sum, s) => sum + (Number(s.totalCost) || 0), 0);
  const basketAvgCost = basketTotalQty > 0 ? basketTotalValue / basketTotalQty : 0;
  // หน่วยของสินค้าที่เลือกอาจไม่เหมือนกัน เช่น กก. vs ชิ้น — เช็คก่อนว่าจำนวนรวมกันได้มีความหมายไหม
  const basketUnits = [...new Set(basketItems.map((s) => s.unit))];
  const basketSameUnit = basketUnits.length <= 1;

  return (
    <div>
      {(totalOpeningStockValue > 0 || totalOpeningBankBalance > 0) && (
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 14 }}>
          {totalOpeningStockValue > 0 && (
            <div style={{ background: "#e1f5ee", border: "1px solid #a3d9c3", borderRadius: 8, padding: "7px 14px", fontSize: 13, color: "#0c443c", display: "flex", gap: 6, alignItems: "center" }}>
              <Boxes size={14} />
              <span>สต็อกยกมา <strong>{fmt(totalOpeningStockQty)} หน่วย</strong> มูลค่า <strong>฿{fmt(totalOpeningStockValue)}</strong> — รวมในสต็อกแล้ว</span>
            </div>
          )}
          {totalOpeningBankBalance > 0 && (
            <div style={{ background: "#e6f1fb", border: "1px solid #b3d0f0", borderRadius: 8, padding: "7px 14px", fontSize: 13, color: "#0c447c", display: "flex", gap: 6, alignItems: "center" }}>
              <Landmark size={14} />
              <span>ยอดธนาคารยกมารวม <strong>฿{fmt(totalOpeningBankBalance)}</strong> ({(storeBankAccounts || []).filter(a => Number(a.openingBalance) > 0).length} บัญชี)</span>
            </div>
          )}
        </div>
      )}
      <Header title="ระบบสต๊อกสินค้า (Inventory)" subtitle="ติดตามยอดรับเข้า-เบิกออก คำนวณต้นทุนแบบ FIFO">
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          {selectedIds.size > 0 && (
            <button style={btnSecondary} onClick={() => setShowBasket(true)}>
              <ShoppingCart size={16} /> ดูตะกร้า
              <span style={{ background: "#534ab7", color: "#fff", borderRadius: 999, fontSize: 11, fontWeight: 700, padding: "1px 7px", marginLeft: 4 }}>{selectedIds.size}</span>
            </button>
          )}
          <ExportToolbar
            onPDF={() => printAsPDF("tab-export-inventory", "สต๊อกสินค้า")}
            onExcel={() => {
              const rows = [
                ["สินค้า", "คงเหลือ", "หน่วย", "ต้นทุนเฉลี่ย/หน่วย", "มูลค่าคงเหลือ"],
                ...inventory.summary.map((s) => [s.name, s.qty, s.unit, s.avgCost, s.totalCost]),
              ];
              exportExcel(rows, "สต๊อกสินค้า.xlsx", "สต๊อก");
            }}
            onImage={() => printAsPDF("tab-export-inventory", "สต๊อกสินค้า")}
          />
        </div>
      </Header>
      <div id="tab-export-inventory">
      <Card style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ overflowY: "auto", maxHeight: "calc(100vh - 260px)" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead style={{ position: "sticky", top: 0, zIndex: 10, background: "#f9fafb" }}>
            <tr>
              <th style={thStyle}></th>
              <th style={thStyle}>สินค้า</th>
              <th style={{ ...thStyle, textAlign: "right" }}>คงเหลือ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ต้นทุนเฉลี่ย/หน่วย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>มูลค่าคงเหลือ</th>
              <th style={{ ...thStyle, textAlign: "center" }}>เลือก</th>
            </tr>
          </thead>
          <tbody>
            {inventory.summary.map((s) => (
              <React.Fragment key={s.productId}>
                <tr style={{ cursor: "pointer", background: selectedIds.has(s.productId) ? "#f3f1fd" : "transparent" }} onClick={() => setExpanded(expanded === s.productId ? null : s.productId)}>
                  <td style={tdStyle}>{expanded === s.productId ? <ChevronDown size={16} /> : <ChevronRight size={16} />}</td>
                  <td style={{ ...tdStyle, fontWeight: 500 }}>{s.name}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{fmtInt(s.qty)} {s.unit}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(s.avgCost)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>{fmt(s.totalCost)}</td>
                  <td style={{ ...tdStyle, textAlign: "center" }} onClick={(e) => e.stopPropagation()}>
                    <input type="checkbox" checked={selectedIds.has(s.productId)} onChange={() => toggleSelect(s.productId)} style={{ width: 16, height: 16, cursor: "pointer" }} />
                  </td>
                </tr>
                {expanded === s.productId && (
                  <tr>
                    <td colSpan={6} style={{ padding: "0 12px 16px 36px", background: "#f9fafb" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, fontWeight: 600, margin: "12px 0 8px", color: "#374151" }}>
                        <History size={14} /> ประวัติการเคลื่อนไหวสินค้า
                      </div>
                      {(inventory.history[s.productId] || []).length === 0 ? (
                        <p style={{ color: "#9ca3af", fontSize: 13 }}>ยังไม่มีการเคลื่อนไหว</p>
                      ) : (
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                          <thead>
                            <tr>
                              <th style={thStyle}>วันที่</th>
                              <th style={thStyle}>เลขที่อ้างอิง</th>
                              <th style={thStyle}>ประเภท</th>
                              <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                              <th style={{ ...thStyle, textAlign: "right" }}>คงเหลือสะสม</th>
                              <th style={{ ...thStyle, textAlign: "right" }}>ราคา/ต้นทุน</th>
                            </tr>
                          </thead>
                          <tbody>
                            {inventory.history[s.productId].map((ev, idx) => (
                              <tr key={idx}>
                                <td style={tdStyle}>{ev.date}</td>
                                <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace" }}>{ev.ref}</td>
                                <td style={tdStyle}>
                                  {ev.type === "in" ? (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, color: "#0f6e56" }}><ArrowDownToLine size={14} /> รับเข้า</span>
                                  ) : ev.type === "withdraw" ? (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, color: "#534ab7" }}><PackageMinus size={14} /> เบิกเพื่อขาย</span>
                                  ) : (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, color: "#993c1d" }}><ArrowUpFromLine size={14} /> เบิกออก</span>
                                  )}
                                </td>
                                <td style={{ ...tdStyle, textAlign: "right", color: ev.type === "in" ? "#0f6e56" : ev.type === "withdraw" ? "#534ab7" : "#993c1d" }}>
                                  {ev.type === "in" ? "+" : "-"}{fmt(ev.qty)} {s.unit}
                                </td>
                                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 500 }}>{fmt(ev.balance)} {s.unit}</td>
                                <td style={{ ...tdStyle, textAlign: "right" }}>
                                  {ev.type === "in" ? fmt(ev.price) : fmt(ev.avgCostUsed)}
                                  {ev.type !== "in" && ev.shortfall > 0 && (
                                    <span style={{ color: "#a32d2d", marginLeft: 6, fontSize: 11 }}>(ขาด {fmt(ev.shortfall)})</span>
                                  )}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      )}
                    </td>
                  </tr>
                )}
              </React.Fragment>
            ))}
          </tbody>
          {inventory.summary.length > 0 && (
            <tfoot>
              <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                <td style={tdStyle}></td>
                <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด ({inventory.summary.length} รายการ)</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>—</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>—</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#185fa5" }}>฿{fmt(inventory.summary.reduce((s, x) => s + x.totalCost, 0))}</td>
                <td style={tdStyle}></td>
              </tr>
            </tfoot>
          )}
        </table>
        </div>{/* end scroll wrapper */}
      </Card>
      </div>{/* end tab-export-inventory */}

      {showBasket && (
        <Modal title={`ตะกร้าสินค้าที่เลือก (${basketItems.length} รายการ)`} onClose={() => setShowBasket(false)} wide>
          {basketItems.length === 0 ? (
            <p style={{ color: "#9ca3af", fontSize: 13, textAlign: "center", padding: 20 }}>ยังไม่ได้เลือกสินค้า</p>
          ) : (
            <>
              <div id="basket-pdf-content">
                <h3 style={{ margin: "0 0 12px", fontSize: 15, fontWeight: 700 }}>ตะกร้าสินค้าที่เลือก — {today}</h3>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, minWidth: 480 }}>
                    <thead>
                      <tr>
                        <th style={thStyle}>สินค้า</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>ราคาเฉลี่ย/หน่วย</th>
                        <th style={{ ...thStyle, textAlign: "right" }}>มูลค่า</th>
                        <th style={thStyle}></th>
                      </tr>
                    </thead>
                    <tbody>
                      {basketItems.map((s) => (
                        <tr key={s.productId}>
                          <td style={tdStyle}>{s.name}</td>
                          <td style={{ ...tdStyle, textAlign: "right" }}>{fmtInt(s.qty)} {s.unit}</td>
                          <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(s.avgCost)}</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(s.totalCost)}</td>
                          <td style={{ ...tdStyle, textAlign: "center" }}>
                            <button style={iconBtn} onClick={() => toggleSelect(s.productId)} aria-label="นำออกจากตะกร้า"><X size={14} /></button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", marginTop: 14, fontSize: 14 }}>
                  {basketSameUnit ? (
                    <Row label="จำนวนรวม" value={`${fmt(basketTotalQty)} ${basketUnits[0] || ""}`} />
                  ) : (
                    <p style={{ fontSize: 12, color: "#9ca3af", margin: "0 0 8px" }}>* สินค้าที่เลือกมีหน่วยไม่เหมือนกัน จึงไม่รวมจำนวนให้</p>
                  )}
                  <Row label="ราคาเฉลี่ยถ่วงน้ำหนัก" value={basketSameUnit ? `฿${fmt(basketAvgCost)}` : "-"} />
                  <Row label="ยอดรวมมูลค่า" value={`฿${fmt(basketTotalValue)}`} bold color="#534ab7" />
                </div>
              </div>

              <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
                <button style={btnSecondary} onClick={clearBasket}><Trash2 size={14} /> ล้างตะกร้า</button>
                <button style={btnSecondary} onClick={() => printAsPDF("basket-pdf-content", "ตะกร้าสินค้าที่เลือก")}><Download size={14} /> พิมพ์ / PDF</button>
                <button style={btnPrimary} onClick={() => setShowBasket(false)}>ปิด</button>
              </div>
            </>
          )}
        </Modal>
      )}
    </div>
  );
}
// DEPOSITS TAB (เงินมัดจำจ่ายล่วงหน้าให้ลูกค้า)
// ===================================================================
function DepositsTab({ customers, setCustomers, deposits, setDeposits, purchases, storeBankAccounts }) {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [openingModal, setOpeningModal] = useState(null); // {customerId, amount}

  const custName = (id) => customers.find((c) => c.id === id)?.name || id;

  const blankForm = () => ({
    id: genId("AE", deposits),
    date: new Date().toISOString().slice(0, 10),
    customerId: "",
    amount: 0,
    fromStoreBankId: storeBankAccounts[0]?.id || "",
    note: "",
  });
  const [form, setForm] = useState(blankForm());

  const openAdd = () => { setForm(blankForm()); setModal({ mode: "add" }); };
  const openEdit = (item) => { setForm({ ...item }); setModal({ mode: "edit", item }); };

  const save = () => {
    if (!form.customerId || !(Number(form.amount) > 0)) return;
    const cleaned = { ...form, amount: Number(form.amount) || 0 };
    if (modal.mode === "add") setDeposits([...deposits, cleaned]);
    else setDeposits(deposits.map((d) => (d.id === modal.item.id ? cleaned : d)));
    setModal(null);
  };

  const remove = (id) => setDeposits(deposits.filter((d) => d.id !== id));

  const balances = useMemo(() => computeDepositBalances(customers, deposits, purchases), [customers, deposits, purchases]);

  // รายการหักมัดจำที่เกิดขึ้นในใบรับสินค้าทั้งหมด (สำหรับแสดงประวัติการใช้)
  const depositUsages = useMemo(() => {
    const list = [];
    purchases.forEach((po) => {
      (po.payments || []).forEach((p) => {
        if (p.fromStoreBankId === "DEPOSIT") {
          list.push({ id: `${po.id}-${p.id}`, date: p.date, customerId: po.customerId, poId: po.id, amount: Number(p.amount) || 0 });
        }
      });
    });
    return list;
  }, [purchases]);

  const filtered = deposits.filter((d) => custName(d.customerId).includes(search) || d.id.includes(search)).filter((d) => (!dateFrom || (d.date || "") >= dateFrom) && (!dateTo || (d.date || "") <= dateTo)).sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  const fromLabel = (id) => {
    if (id === "CASH") return "เงินสดหน้าร้าน";
    const b = storeBankAccounts.find((b) => b.id === id);
    return b ? `${b.bankName} ${b.accountNo}` : "-";
  };

  const openOpeningModal = () => setOpeningModal({ customerId: customers[0]?.id || "", amount: "" });
  const editOpeningModal = (customerId, currentValue) => setOpeningModal({ customerId, amount: String(currentValue || 0) });
  const saveOpening = () => {
    if (!openingModal || !openingModal.customerId) return;
    setCustomers(customers.map((c) => c.id === openingModal.customerId ? { ...c, depositOpening: Number(openingModal.amount) || 0 } : c));
    setOpeningModal(null);
  };

  return (
    <div>
      <Header title="เงินมัดจำ (จ่ายล่วงหน้าให้ลูกค้า)" subtitle="บันทึกเงินมัดจำที่จ่ายให้ลูกค้าล่วงหน้า และดูยอดมัดจำคงเหลือของลูกค้าแต่ละราย">
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button style={btnSecondary} onClick={openOpeningModal}><Plus size={16} /> เพิ่มยอดยกมา</button>
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> บันทึกจ่ายมัดจำ</button>
        </div>
      </Header>

      <div style={{ marginBottom: 20 }}>
        <h3 style={{ fontSize: 15, fontWeight: 600, margin: "0 0 10px" }}>สรุปยอดมัดจำคงเหลือต่อลูกค้า</h3>
        <Card>
          <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 680 }}>
            <thead>
              <tr>
                <th style={thStyle}>ลูกค้า</th>
                <th style={{ ...thStyle, textAlign: "right" }}>ยอดยกมา</th>
                <th style={{ ...thStyle, textAlign: "right" }}>มัดจำที่จ่ายใหม่</th>
                <th style={{ ...thStyle, textAlign: "right" }}>มัดจำที่จ่ายรวม</th>
                <th style={{ ...thStyle, textAlign: "right" }}>หักไปแล้ว (ในใบรับสินค้า)</th>
                <th style={{ ...thStyle, textAlign: "right" }}>คงเหลือ</th>
              </tr>
            </thead>
            <tbody>
              {balances.filter((b) => b.totalGiven > 0 || b.opening > 0).map((b) => (
                <tr key={b.customerId}>
                  <td style={tdStyle}>{b.name}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    <span style={{ cursor: "pointer", borderBottom: "1px dashed #9ca3af" }} onClick={() => editOpeningModal(b.customerId, b.opening)} title="คลิกเพื่อแก้ไขยอดยกมา">
                      ฿{fmt(b.opening)}
                    </span>
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(b.newGiven)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(b.totalGiven)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", color: "#854f0b" }}>฿{fmt(b.totalUsed)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: b.remaining > 0 ? "#0f6e56" : "#6b7280" }}>฿{fmt(b.remaining)}</td>
                </tr>
              ))}
              {balances.every((b) => b.totalGiven === 0 && b.opening === 0) && (
                <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีการจ่ายมัดจำ — กดปุ่ม "เพิ่มยอดยกมา" หรือ "บันทึกจ่ายมัดจำ" ด้านบนเพื่อเริ่มต้น</td></tr>
              )}
            </tbody>
          </table>
          </div>
          <p style={{ fontSize: 12, color: "#9ca3af", margin: "8px 0 0", padding: "0 4px" }}>* คลิกที่ตัวเลข "ยอดยกมา" เพื่อแก้ไขได้โดยตรง</p>
        </Card>
      </div>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาลูกค้า หรือเลขที่รายการมัดจำ..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />

      <h3 style={{ fontSize: 15, fontWeight: 600, margin: "0 0 10px" }}>ประวัติการจ่ายมัดจำ</h3>
      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={thStyle}>วันที่</th>
              <th style={thStyle}>ลูกค้า</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
              <th style={thStyle}>จ่ายจาก</th>
              <th style={thStyle}>หมายเหตุ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((d) => (
              <tr key={d.id}>
                <td style={tdStyle}>{d.date}</td>
                <td style={tdStyle}>{custName(d.customerId)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#0f6e56" }}>+฿{fmt(d.amount)}</td>
                <td style={tdStyle}>{fromLabel(d.fromStoreBankId)}</td>
                <td style={{ ...tdStyle, color: "#6b7280" }}>{d.note || "-"}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <button style={iconBtn} onClick={() => openEdit(d)}><Edit2 size={14} /> แก้ไข</button>
                    <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบรายการมัดจำของ "${custName(d.customerId)}" จำนวน ฿${fmt(d.amount)} ใช่หรือไม่?`, () => remove(d.id))}><Trash2 size={14} /> ลบ</button>
                  </div>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีรายการมัดจำ</td></tr>}
          </tbody>
          {filtered.length > 0 && (
            <tfoot>
              <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                <td style={{ ...tdStyle, fontWeight: 700 }} colSpan={2}>รวมทั้งหมด ({filtered.length} รายการ)</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>+฿{fmt(filtered.reduce((s, d) => s + (Number(d.amount) || 0), 0))}</td>
                <td colSpan={3} style={tdStyle}></td>
              </tr>
            </tfoot>
          )}
        </table>
      </Card>

      {depositUsages.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <h3 style={{ fontSize: 15, fontWeight: 600, margin: "0 0 10px" }}>ประวัติการหักมัดจำ (จากใบรับสินค้า)</h3>
          <Card>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={thStyle}>วันที่</th>
                  <th style={thStyle}>ลูกค้า</th>
                  <th style={thStyle}>เลขที่ใบรับสินค้า</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงินที่หัก</th>
                </tr>
              </thead>
              <tbody>
                {depositUsages.map((u) => (
                  <tr key={u.id}>
                    <td style={tdStyle}>{u.date}</td>
                    <td style={tdStyle}>{custName(u.customerId)}</td>
                    <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace" }}>{u.poId}</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#993c1d" }}>-฿{fmt(u.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>
        </div>
      )}

      {modal && (
        <Modal title={modal.mode === "add" ? "บันทึกจ่ายมัดจำให้ลูกค้า" : "แก้ไขรายการมัดจำ"} onClose={() => setModal(null)}>
          <Field label="วันที่จ่าย">
            <input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
          </Field>
          <Field label="ลูกค้า">
            <CustomerSelect customers={customers} value={form.customerId} onChange={(cid) => setForm({ ...form, customerId: cid })} />
          </Field>
          <Field label="จำนวนเงินมัดจำ">
            <input type="number" style={inputStyle} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} />
          </Field>
          <Field label="จ่ายจาก">
            <select style={inputStyle} value={form.fromStoreBankId} onChange={(e) => setForm({ ...form, fromStoreBankId: e.target.value })}>
              <option value="">-- เลือกบัญชี --</option>
              {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
            </select>
          </Field>
          <Field label="หมายเหตุ">
            <input style={inputStyle} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="เช่น มัดจำสำหรับงานเดือน มิ.ย." />
          </Field>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {openingModal && (
        <Modal title="เพิ่ม/แก้ไขยอดยกมา" onClose={() => setOpeningModal(null)}>
          <Field label="ลูกค้า">
            <CustomerSelect customers={customers} value={openingModal.customerId} onChange={(cid) => setOpeningModal({ ...openingModal, customerId: cid, amount: String(customers.find((c) => c.id === cid)?.depositOpening || 0) })} />
          </Field>
          <Field label="ยอดยกมา (มัดจำคงเหลือเดิม)">
            <input type="number" style={inputStyle} value={openingModal.amount} onChange={(e) => setOpeningModal({ ...openingModal, amount: e.target.value })} placeholder="0" />
          </Field>
          <p style={{ fontSize: 12, color: "#9ca3af", margin: "-4px 0 8px" }}>
            * ยอดนี้คือมัดจำคงเหลือเดิมก่อนเริ่มใช้ระบบ จะรวมเข้ากับมัดจำที่จ่ายใหม่และใช้หักในใบรับสินค้าได้จริง
          </p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setOpeningModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={saveOpening}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ===================================================================
// LOANS TAB (เงินกู้ยืม / เช่าซื้อ)
// ===================================================================
function LoansTab({ loans, setLoans, expenses, customers }) {
  const [modal, setModal] = useState(null); // {mode:'add'|'edit'|'schedule', item}

  const blankForm = () => ({
    id: genId("CT", loans),
    billNo: "",
    name: "",
    type: LOAN_TYPES[0],
    lenderCustomerId: "", // อ้างอิงลูกค้าจากฐานข้อมูล (ถ้ามี)
    lender: "", // ชื่อผู้ให้กู้/ไฟแนนซ์ (พิมพ์เองได้ ถ้าไม่มีในฐานข้อมูลลูกค้า)
    principal: 0,
    interestMode: "rate", // "rate" = % ต่อปี, "amount" = จำนวนเงินดอกเบี้ยรวมตลอดสัญญา
    annualInterestRate: 0,
    totalInterestAmount: 0,
    totalInstallments: 12,
    startDate: new Date().toISOString().slice(0, 10),
    dueDayOfMonth: new Date().getDate(), // ครบกำหนดชำระทุกวันที่เท่าไรของเดือน (1-31)
    paidInstallments: [], // [{no, expenseId, paidDate}]
  });
  const [form, setForm] = useState(blankForm());

  const openAdd = () => { setForm(blankForm()); setModal({ mode: "add" }); };
  const openEdit = (item) => { setForm(JSON.parse(JSON.stringify({ paidInstallments: [], interestMode: "rate", totalInterestAmount: 0, lenderCustomerId: "", billNo: "", dueDayOfMonth: new Date(item.startDate || Date.now()).getDate(), ...item }))); setModal({ mode: "edit", item }); };
  const openSchedule = (item) => setModal({ mode: "schedule", item });

  const save = () => {
    if (!form.name.trim() || !(Number(form.principal) > 0) || !(Number(form.totalInstallments) > 0)) return;
    const cleaned = {
      ...form,
      principal: Number(form.principal) || 0,
      annualInterestRate: Number(form.annualInterestRate) || 0,
      totalInterestAmount: Number(form.totalInterestAmount) || 0,
      totalInstallments: Number(form.totalInstallments) || 0,
    };
    if (modal.mode === "add") setLoans([...loans, cleaned]);
    else setLoans(loans.map((l) => (l.id === modal.item.id ? cleaned : l)));
    setModal(null);
  };

  const remove = (id) => setLoans(loans.filter((l) => l.id !== id));

  // เมื่อเลือกผู้ให้กู้จากฐานข้อมูลลูกค้า ให้เติมชื่อลงในช่อง lender ด้วย (ใช้แสดงผล/ค้นหาในตาราง)
  const handleLenderChange = (customerId) => {
    const c = customers.find((c) => c.id === customerId);
    setForm({ ...form, lenderCustomerId: customerId, lender: c ? c.name : form.lender });
  };

  const preview = useMemo(() => computeAmortizationSchedule(form), [form.principal, form.annualInterestRate, form.totalInterestAmount, form.interestMode, form.totalInstallments, form.startDate, form.dueDayOfMonth]);

  return (
    <div>
      <Header title="เงินกู้ยืม / เช่าซื้อ" subtitle="บันทึกสัญญาเงินกู้/เช่าซื้อ และตารางผ่อนชำระ สามารถดึงงวดมาตัดจ่ายในหน้าค่าใช้จ่ายได้">
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <ExportToolbar
            onPDF={() => printAsPDF("tab-export-loans", "เงินกู้ยืม/เช่าซื้อ")}
            onExcel={() => {
              const rows = [
                ["ชื่อสัญญา", "เลขที่บิล/สัญญา", "ประเภท", "เงินต้น", "ดอกเบี้ย", "จำนวนงวด", "ผ่อนแล้ว", "คงเหลือ (งวด)"],
                ...loans.map((l) => {
                  const paidCount = (l.paidInstallments || []).length;
                  return [l.name, l.billNo || "", l.type, l.principal, l.interestMode === "amount" ? l.totalInterestAmount : `${l.annualInterestRate}%/ปี`, l.totalInstallments, paidCount, l.totalInstallments - paidCount];
                }),
              ];
              exportExcel(rows, "เงินกู้เช่าซื้อ.xlsx", "สินเชื่อ");
            }}
            onImage={() => printAsPDF("tab-export-loans", "เงินกู้ยืม/เช่าซื้อ")}
          />
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> เพิ่มสัญญาเงินกู้/เช่าซื้อ</button>
        </div>
      </Header>
      <div id="tab-export-loans">

      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={thStyle}>ชื่อสัญญา</th>
              <th style={thStyle}>เลขที่บิล/สัญญา</th>
              <th style={thStyle}>ประเภท</th>
              <th style={thStyle}>ผู้ให้กู้/ไฟแนนซ์</th>
              <th style={{ ...thStyle, textAlign: "right" }}>เงินต้น</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ดอกเบี้ย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จำนวนงวด</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ผ่อนแล้ว</th>
              <th style={{ ...thStyle, textAlign: "right" }}>คงเหลือ</th>
              <th style={thStyle}>วันครบกำหนดชำระ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {[...loans].sort((a, b) => (b.startDate || "").localeCompare(a.startDate || "") || (b.id || "").localeCompare(a.id || "")).map((l) => {
              const schedule = computeAmortizationSchedule(l);
              const paidCount = (l.paidInstallments || []).length;
              const remainingCount = l.totalInstallments - paidCount;
              // วันครบกำหนดของงวดถัดไป = งวดที่ (จำนวนงวดที่ชำระแล้ว + 1) ตามลำดับ ต่อจากงวดที่ชำระแล้ว
              const nextInstallment = schedule.find((s) => s.no === paidCount + 1);
              const nextDueDate = nextInstallment?.dueDate;
              return (
                <tr key={l.id}>
                  <td style={tdStyle}>{l.name}</td>
                  <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{l.billNo || "-"}</td>
                  <td style={tdStyle}><Badge text={l.type} /></td>
                  <td style={tdStyle}>{l.lender || "-"}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(l.principal)}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    {l.interestMode === "amount" ? `฿${fmt(l.totalInterestAmount)} (รวม)` : `${fmt(l.annualInterestRate)}% /ปี`}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{l.totalInstallments} งวด</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: paidCount >= l.totalInstallments ? "#0f6e56" : "#854f0b" }}>
                    {paidCount} / {l.totalInstallments}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: remainingCount > 0 ? "#993c1d" : "#0f6e56" }}>
                    {remainingCount} งวด
                  </td>
                  <td style={tdStyle}>
                    <div>ทุกวันที่ {l.dueDayOfMonth || "-"} ของเดือน</div>
                    {nextDueDate && <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>งวดถัดไป: {nextDueDate}</div>}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>
                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <button style={iconBtn} onClick={() => openSchedule(l)}><History size={14} /> ตารางผ่อน</button>
                      <button style={iconBtn} onClick={() => openEdit(l)}><Edit2 size={14} /> แก้ไข</button>
                      <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบสัญญา "${l.name || l.id}" ใช่หรือไม่?`, () => remove(l.id))}><Trash2 size={14} /> ลบ</button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {loans.length === 0 && <tr><td colSpan={11} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีสัญญาเงินกู้/เช่าซื้อ</td></tr>}
          </tbody>
        </table>
      </Card>

      {modal && (modal.mode === "add" || modal.mode === "edit") && (
        <Modal title={modal.mode === "add" ? "เพิ่มสัญญาเงินกู้/เช่าซื้อ" : `แก้ไขสัญญา · ${form.name}`} onClose={() => setModal(null)} wide>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 16px" }}>
            <Field label="ชื่อสัญญา">
              <input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="เช่น สินเชื่อรถบรรทุก 6 ล้อ" />
            </Field>
            <Field label="เลขที่บิล/สัญญา">
              <input style={inputStyle} value={form.billNo} onChange={(e) => setForm({ ...form, billNo: e.target.value })} placeholder="เช่น CT-2026-0001" />
            </Field>
            <Field label="ประเภท">
              <select style={inputStyle} value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                {LOAN_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </Field>
          </div>
          <Field label="ผู้ให้กู้ / ไฟแนนซ์ (ค้นหาจากฐานข้อมูลลูกค้า)">
            <CustomerSelect customers={customers} value={form.lenderCustomerId} onChange={handleLenderChange} labelWithId={false} />
          </Field>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: "0 16px" }}>
            <Field label="เงินต้น (บาท)">
              <input type="number" style={inputStyle} value={form.principal} onChange={(e) => setForm({ ...form, principal: e.target.value })} />
            </Field>
            <Field label="รูปแบบดอกเบี้ย">
              <select style={inputStyle} value={form.interestMode} onChange={(e) => setForm({ ...form, interestMode: e.target.value })}>
                <option value="rate">% ต่อปี (ลดต้นลดดอก)</option>
                <option value="amount">กรอกจำนวนเงินดอกเบี้ยรวม</option>
              </select>
            </Field>
            {form.interestMode === "amount" ? (
              <Field label="ดอกเบี้ยรวมตลอดสัญญา (บาท)">
                <input type="number" style={inputStyle} value={form.totalInterestAmount} onChange={(e) => setForm({ ...form, totalInterestAmount: e.target.value })} />
              </Field>
            ) : (
              <Field label="ดอกเบี้ย (% ต่อปี)">
                <input type="number" style={inputStyle} value={form.annualInterestRate} onChange={(e) => setForm({ ...form, annualInterestRate: e.target.value })} />
              </Field>
            )}
            <Field label="จำนวนงวด (เดือน)">
              <input type="number" style={inputStyle} value={form.totalInstallments} onChange={(e) => setForm({ ...form, totalInstallments: e.target.value })} />
            </Field>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            <Field label="วันที่เริ่มสัญญา">
              <input type="date" style={inputStyle} value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} />
            </Field>
            <Field label="ครบกำหนดชำระทุกวันที่ (1-31) ของเดือน">
              <input
                type="number"
                min={1}
                max={31}
                style={inputStyle}
                value={form.dueDayOfMonth}
                onChange={(e) => setForm({ ...form, dueDayOfMonth: e.target.value })}
              />
              <p style={{ fontSize: 12, color: "#9ca3af", marginTop: 4, marginBottom: 0 }}>
                * ถ้าวันที่เกินจำนวนวันในเดือนนั้น (เช่น 31 ในเดือน ก.พ.) ระบบจะใช้วันสุดท้ายของเดือนแทน
              </p>
            </Field>
          </div>

          {preview.length > 0 && (
            <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", marginBottom: 16, fontSize: 14 }}>
              <Row label="ค่างวดผ่อนต่อเดือน (งวดแรก)" value={`฿${fmt(preview[0].payment)}`} bold />
              <Row label="ดอกเบี้ยรวมตลอดสัญญา" value={`฿${fmt(preview.reduce((s, p) => s + p.interest, 0))}`} />
              <Row label="ยอดชำระรวมตลอดสัญญา" value={`฿${fmt(preview.reduce((s, p) => s + p.payment, 0))}`} />
            </div>
          )}

          {preview.length > 0 && (
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 8 }}>สรุปยอดจ่ายต่องวด</div>
              <div style={{ overflowX: "auto", maxHeight: 240, overflowY: "auto", border: "1px solid #e5e7eb", borderRadius: 8 }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr>
                      <th style={thStyle}>งวดที่</th>
                      <th style={thStyle}>วันครบกำหนด</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>ยอดผ่อน</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>ดอกเบี้ย</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>ตัดเงินต้น</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>เงินต้นคงเหลือ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {preview.map((p) => (
                      <tr key={p.no}>
                        <td style={tdStyle}>{p.no}</td>
                        <td style={tdStyle}>{p.dueDate}</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(p.payment)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#854f0b" }}>฿{fmt(p.interest)}</td>
                        <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(p.principalPortion)}</td>
                        <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(p.remainingBalance)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {modal && modal.mode === "schedule" && (
        <LoanScheduleModal loan={modal.item} expenses={expenses} onClose={() => setModal(null)} />
      )}
      </div>{/* end tab-export-loans */}
    </div>
  );
}

// ตารางผ่อนชำระแบบละเอียด พร้อมสถานะการจ่าย
function LoanScheduleModal({ loan, expenses, onClose }) {
  const schedule = computeAmortizationSchedule(loan);
  const paidMap = {};
  (loan.paidInstallments || []).forEach((p) => { paidMap[p.no] = p; });

  return (
    <Modal title={`ตารางผ่อนชำระ · ${loan.name}${loan.billNo ? " · " + loan.billNo : ""}`} onClose={onClose} wide>
      <div style={{ background: "#f9fafb", borderRadius: 8, padding: "10px 16px", marginBottom: 12, fontSize: 13 }}>
        <Row label="เงินต้น" value={`฿${fmt(loan.principal)}`} />
        <Row label="ดอกเบี้ย" value={loan.interestMode === "amount" ? `฿${fmt(loan.totalInterestAmount)} (รวม)` : `${fmt(loan.annualInterestRate)}% ต่อปี`} />
        <Row label="ค่างวดต่อเดือน" value={schedule.length > 0 ? `฿${fmt(schedule[0].payment)}` : "-"} bold />
      </div>
      <div style={{ overflowX: "auto", maxHeight: 420, overflowY: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr>
              <th style={thStyle}>งวดที่</th>
              <th style={thStyle}>วันครบกำหนด</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ยอดผ่อน</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ดอกเบี้ย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ตัดเงินต้น</th>
              <th style={{ ...thStyle, textAlign: "right" }}>เงินต้นคงเหลือ</th>
              <th style={thStyle}>สถานะ</th>
            </tr>
          </thead>
          <tbody>
            {schedule.map((s) => {
              const paid = paidMap[s.no];
              return (
                <tr key={s.no} style={paid ? { background: "#e1f5ee" } : undefined}>
                  <td style={tdStyle}>{s.no}</td>
                  <td style={tdStyle}>{s.dueDate}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(s.payment)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", color: "#854f0b" }}>฿{fmt(s.interest)}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(s.principalPortion)}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(s.remainingBalance)}</td>
                  <td style={tdStyle}>
                    {paid ? (
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4, color: "#0f6e56", fontWeight: 600 }}><CheckCircle2 size={14} /> จ่ายแล้ว ({paid.paidDate})</span>
                    ) : (
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4, color: "#9ca3af" }}><Clock size={14} /> ยังไม่จ่าย</span>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
        <button style={btnSecondary} onClick={onClose}>ปิด</button>
      </div>
    </Modal>
  );
}

// ===================================================================
// EXPENSES TAB (บันทึกค่าใช้จ่าย / ใบสำคัญจ่าย)
// ===================================================================
// ===================================================================
// PREPAYMENTS TAB (รับล่วงหน้า — ลูกค้าจ่ายให้ร้านก่อนรับสินค้า)
// ===================================================================
function PrepaymentsTab({ customers, setCustomers, prepayments, setPrepayments, sales, storeBankAccounts }) {
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [modal, setModal] = useState(null);
  const [openingModal, setOpeningModal] = useState(null);
  const [form, setForm] = useState({ id: "", date: new Date().toISOString().slice(0, 10), customerId: customers[0]?.id || "", amount: 0, toStoreBankId: storeBankAccounts[0]?.id || "", note: "" });

  const openOpeningModal = () => setOpeningModal({ customerId: customers[0]?.id || "", amount: "" });

  const balances = useMemo(() => computePrepaymentBalances(customers, prepayments, sales), [customers, prepayments, sales]);
  const custName = (id) => customers.find((c) => c.id === id)?.name || id;
  const bankName = (id) => { const b = storeBankAccounts.find((b) => b.id === id); return b ? `${b.bankName} ${b.accountNo}` : "-"; };
  const fmt = (n) => Number(n).toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  const filtered = prepayments
    .filter((d) => custName(d.customerId).includes(search) || d.id.includes(search))
    .filter((d) => (!dateFrom || (d.date || "") >= dateFrom) && (!dateTo || (d.date || "") <= dateTo))
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  const openAdd = () => {
    const id = "PP" + Date.now().toString().slice(-8);
    setForm({ id, date: new Date().toISOString().slice(0, 10), customerId: customers[0]?.id || "", amount: 0, toStoreBankId: storeBankAccounts[0]?.id || "", note: "" });
    setModal({ mode: "add" });
  };

  const save = () => {
    const cleaned = { ...form, amount: Number(form.amount) || 0 };
    if (modal.mode === "add") setPrepayments([...prepayments, cleaned]);
    else setPrepayments(prepayments.map((d) => d.id === modal.item.id ? cleaned : d));
    setModal(null);
  };

  const remove = (id) => setPrepayments(prepayments.filter((d) => d.id !== id));

  return (
    <div style={{ padding: 24 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <div>
          <h2 style={{ fontWeight: 700, fontSize: 20, margin: 0 }}>รับล่วงหน้า</h2>
          <div style={{ color: "#6b7280", fontSize: 13 }}>บันทึกเงินที่ลูกค้าจ่ายล่วงหน้าก่อนรับสินค้า</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={btnSecondary} onClick={openOpeningModal}><Plus size={16} /> เพิ่มยอดยกมา</button>
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> บันทึกรับล่วงหน้า</button>
        </div>
      </div>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาลูกค้า..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />

      {/* สรุปยอดคงเหลือต่อลูกค้า */}
      <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "auto", marginBottom: 20 }}>
        <div style={{ padding: "14px 20px", borderBottom: "1px solid #f3f4f6", fontWeight: 600, fontSize: 14 }}>สรุปยอดรับล่วงหน้าต่อลูกค้า</div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={thStyle}>ลูกค้า</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ยอดยกมา</th>
              <th style={{ ...thStyle, textAlign: "right" }}>รับเพิ่ม</th>
              <th style={{ ...thStyle, textAlign: "right" }}>รับรวม</th>
              <th style={{ ...thStyle, textAlign: "right" }}>หักไปแล้ว (ในใบขาย)</th>
              <th style={{ ...thStyle, textAlign: "right", color: "#0f6e56" }}>คงเหลือ</th>
            </tr>
          </thead>
          <tbody>
            {balances.filter((b) => b.totalReceived > 0 || b.opening > 0).map((b) => (
              <tr key={b.customerId}>
                <td style={tdStyle}>
                  <button style={{ background: "none", border: "none", color: "#185fa5", cursor: "pointer", fontWeight: 500, padding: 0, textDecoration: "underline", fontSize: 13 }}
                    onClick={() => setOpeningModal({ customerId: b.customerId, amount: String(b.opening) })}>
                    {b.name}
                  </button>
                </td>
                <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280" }}>฿{fmt(b.opening)}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>+฿{fmt(b.newReceived)}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(b.totalReceived)}</td>
                <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d" }}>-฿{fmt(b.totalUsed)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: b.remaining > 0 ? "#0f6e56" : "#991b1b" }}>฿{fmt(b.remaining)}</td>
              </tr>
            ))}
            {balances.filter((b) => b.totalReceived > 0 || b.opening > 0).length === 0 && (
              <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีข้อมูลรับล่วงหน้า</td></tr>
            )}
          </tbody>
          {balances.some((b) => b.totalReceived > 0 || b.opening > 0) && (
            <tfoot>
              <tr style={{ background: "#f0fdf4", borderTop: "2px solid #0f6e56" }}>
                <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>฿{fmt(balances.reduce((s, b) => s + b.opening, 0))}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>+฿{fmt(balances.reduce((s, b) => s + b.newReceived, 0))}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>฿{fmt(balances.reduce((s, b) => s + b.totalReceived, 0))}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>-฿{fmt(balances.reduce((s, b) => s + b.totalUsed, 0))}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>฿{fmt(balances.reduce((s, b) => s + b.remaining, 0))}</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>

      {/* รายการบันทึกรับล่วงหน้า */}
      <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={thStyle}>เลขที่</th>
              <th style={thStyle}>วันที่</th>
              <th style={thStyle}>ลูกค้า</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
              <th style={thStyle}>บัญชีร้าน</th>
              <th style={thStyle}>หมายเหตุ</th>
              <th style={thStyle}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((d) => (
              <tr key={d.id}>
                <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{d.id}</td>
                <td style={tdStyle}>{d.date}</td>
                <td style={{ ...tdStyle, fontWeight: 500 }}>{custName(d.customerId)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#0f6e56" }}>+฿{fmt(d.amount)}</td>
                <td style={tdStyle}>{bankName(d.toStoreBankId)}</td>
                <td style={tdStyle}>{d.note || "-"}</td>
                <td style={tdStyle}>
                  <div style={{ display: "flex", gap: 6 }}>
                    <button style={iconBtn} onClick={() => { setForm({ ...d }); setModal({ mode: "edit", item: d }); }}><Edit2 size={14} /></button>
                    <button style={btnDanger} onClick={() => confirmAction(`ลบรายการรับล่วงหน้า ฿${fmt(d.amount)} ของ ${custName(d.customerId)}?`, () => remove(d.id))}><Trash2 size={14} /></button>
                  </div>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีรายการ</td></tr>}
            {filtered.length > 0 && (
              <tr style={{ background: "#f9fafb", borderTop: "2px solid #e5e7eb" }}>
                <td colSpan={3} style={{ ...tdStyle, fontWeight: 700 }}>รวม ({filtered.length} รายการ)</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>+฿{fmt(filtered.reduce((s, d) => s + (Number(d.amount) || 0), 0))}</td>
                <td colSpan={3} style={tdStyle}></td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Modal บันทึกรับล่วงหน้า */}
      {modal && (
        <Modal title={modal.mode === "add" ? "บันทึกรับล่วงหน้า" : "แก้ไขรับล่วงหน้า"} onClose={() => setModal(null)}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            <Field label="วันที่"><input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
            <Field label="ลูกค้า"><CustomerSelect customers={customers} value={form.customerId} onChange={(cid) => setForm({ ...form, customerId: cid })} labelWithId={false} /></Field>
            <Field label="จำนวนเงินที่รับ (บาท)"><NumInput style={{ ...inputStyle, textAlign: "right" }} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></Field>
            <Field label="บัญชีร้านที่รับเงิน">
              <select style={inputStyle} value={form.toStoreBankId} onChange={(e) => setForm({ ...form, toStoreBankId: e.target.value })}>
                {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
              </select>
            </Field>
          </div>
          <Field label="หมายเหตุ"><input style={inputStyle} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {/* Modal ตั้งยอดยกมา */}
      {openingModal && (
        <Modal title="ตั้งยอดยกมา (รับล่วงหน้า)" onClose={() => setOpeningModal(null)}>
          <Field label="ลูกค้า"><CustomerSelect customers={customers} value={openingModal.customerId} onChange={(cid) => setOpeningModal({ ...openingModal, customerId: cid, amount: String(customers.find((c) => c.id === cid)?.prepaymentOpening || 0) })} labelWithId={false} /></Field>
          <Field label="ยอดยกมา (บาท)"><NumInput style={{ ...inputStyle, textAlign: "right" }} value={openingModal.amount} onChange={(e) => setOpeningModal({ ...openingModal, amount: e.target.value })} /></Field>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setOpeningModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={() => {
              setCustomers(customers.map((c) => c.id === openingModal.customerId ? { ...c, prepaymentOpening: Number(openingModal.amount) || 0 } : c));
              setOpeningModal(null);
            }}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

function ExpensesTab({ expenses, setExpenses, storeBankAccounts, loans, setLoans, expenseCategories, setExpenseCategories, companySettings, customers }) {
  const [modal, setModal] = useState(null); // {mode:'add'|'edit'|'view', item}
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [installmentPicker, setInstallmentPicker] = useState(false); // เปิด picker ดึงงวดผ่อน
  const [expandedExpenses, setExpandedExpenses] = useState({}); // { [expenseId]: bool }
  const toggleExpense = (id) => setExpandedExpenses((prev) => ({ ...prev, [id]: !prev[id] }));
  const [pendingInstallment, setPendingInstallment] = useState(null); // {loanId, no} ของงวดที่เลือกไว้ รอบันทึก
  const [pickerLoanId, setPickerLoanId] = useState(""); // สัญญาที่เลือกใน dropdown ของ picker

  // หมวดหมู่ใหญ่/ย่อย เก็บเป็นฐานข้อมูลแยก (expenseCategories) ใช้ร่วมกันทุกเครื่อง บันทึกถาวรขึ้น Supabase

  const blankItem = () => ({
    id: "EXI" + Date.now().toString().slice(-6) + Math.floor(Math.random() * 1000),
    description: "",
    mainCategory: EXPENSE_MAIN_CATEGORIES[0],
    subCategory: (EXPENSE_SUBCATEGORIES_DEFAULT[EXPENSE_MAIN_CATEGORIES[0]] || [])[0] || "",
    amount: 0,
    vatEnabled: false,
    whtRate: 0,
  });

  const todayStr = new Date().toISOString().slice(0, 10);
  const blankForm = () => ({
    id: "EX" + Date.now().toString().slice(-6),
    refNo: genId("EX", expenses, todayStr),
    recordDate: todayStr,
    taxInvoiceNo: "",
    billDate: todayStr,
    vendorId: "",
    vendorName: "",
    items: [blankItem()], // รายการค่าใช้จ่าย (เพิ่มได้หลายรายการในใบเดียว) — แต่ละรายการมี VAT/หัก ณ ที่จ่ายของตัวเอง
    payments: [],
  });
  const [form, setForm] = useState(blankForm());

  const openAdd = () => { setForm(blankForm()); setPendingInstallment(null); setModal({ mode: "add" }); };
  const openEdit = (item) => {
    // รองรับข้อมูลเดิมที่ยังเป็นรายการเดียว (description/mainCategory/subCategory/amount ที่ระดับบนสุด)
    const items = item.items && item.items.length > 0
      ? item.items.map((it) => ({
          vatEnabled: item.vatEnabled || false, // ของเก่าที่ยังไม่มีต่อรายการ ให้สืบทอดจากระดับบิลเดิม
          whtRate: Number(item.whtRate) || 0,
          ...it,
        }))
      : [{
          id: "EXI" + Date.now().toString().slice(-6),
          description: item.description || "",
          mainCategory: item.mainCategory || item.category || EXPENSE_MAIN_CATEGORIES[0],
          subCategory: item.subCategory || (item.mainCategory ? "" : item.category) || "",
          amount: Number(item.amount) || 0,
          vatEnabled: item.vatEnabled || false,
          whtRate: Number(item.whtRate) || 0,
        }];
    setForm(JSON.parse(JSON.stringify({
      payments: [], taxInvoiceNo: "",
      recordDate: item.recordDate || item.billDate || item.date, refNo: item.refNo || item.id,
      ...item,
      items,
    })));
    setPendingInstallment(item.loanInstallment || null);
    setModal({ mode: "edit", item });
  };
  const openView = (item) => setModal({ mode: "view", item });

  // รวมงวดผ่อนที่ยังไม่จ่ายจากทุกสัญญา (สำหรับ picker "ดึงจากงวดผ่อน")
  const unpaidInstallments = useMemo(() => {
    const list = [];
    (loans || []).forEach((loan) => {
      const schedule = computeAmortizationSchedule(loan);
      const paidNos = new Set((loan.paidInstallments || []).map((p) => p.no));
      schedule.forEach((s) => {
        if (!paidNos.has(s.no)) list.push({ loan, installment: s });
      });
    });
    // จัดกลุ่มตามสัญญา (loan.id) ก่อน แล้วเรียงงวดตามวันครบกำหนดภายในสัญญาเดียวกัน
    return list.sort((a, b) => {
      if (a.loan.id !== b.loan.id) return a.loan.id < b.loan.id ? -1 : 1;
      return a.installment.dueDate < b.installment.dueDate ? -1 : 1;
    });
  }, [loans]);

  // รายชื่อสัญญาที่มีงวดค้างชำระ พร้อมวันครบกำหนดชำระงวดถัดไป (สำหรับรายการให้คลิกเลือกใน picker)
  // วันครบกำหนดของงวดถัดไป = งวดที่ (จำนวนงวดที่ชำระแล้ว + 1) ตามลำดับ ต่อจากงวดที่ชำระแล้ว
  const loansWithUnpaid = useMemo(() => {
    const seen = new Set();
    const result = [];
    unpaidInstallments.forEach(({ loan }) => {
      if (!seen.has(loan.id)) {
        seen.add(loan.id);
        const schedule = computeAmortizationSchedule(loan);
        const paidCount = (loan.paidInstallments || []).length;
        const nextInstallment = schedule.find((s) => s.no === paidCount + 1);
        result.push({ ...loan, nextDueDate: nextInstallment?.dueDate });
      }
    });
    return result;
  }, [unpaidInstallments]);

  // งวดค้างชำระของสัญญาที่เลือก
  const pickerInstallments = useMemo(
    () => unpaidInstallments.filter(({ loan }) => loan.id === pickerLoanId),
    [unpaidInstallments, pickerLoanId]
  );

  // เลือกงวดผ่อนจาก picker -> เติมข้อมูลในฟอร์มค่าใช้จ่ายให้อัตโนมัติ
  const applyInstallment = (loan, installment) => {
    const installmentItem = {
      id: "EXI" + Date.now().toString().slice(-6),
      description: `ผ่อนชำระ ${loan.type} "${loan.name}" งวดที่ ${installment.no}/${loan.totalInstallments} (ดอกเบี้ย ฿${fmt(installment.interest)}, เงินต้น ฿${fmt(installment.principalPortion)})`,
      mainCategory: "สินเชื่อ",
      subCategory: loan.type === "เช่าซื้อ" ? "ชำระค่าเช่าซื้อ" : "ชำระเงินกู้ (เงินต้น)",
      amount: Math.round(installment.payment * 100) / 100,
      vatEnabled: false,
      whtRate: 0,
    };
    setForm({
      ...form,
      items: [installmentItem],
      billDate: installment.dueDate,
    });
    setPendingInstallment({ loanId: loan.id, no: installment.no });
    setInstallmentPicker(false);
  };

  // หมวดหมู่ใหญ่ทั้งหมด: ค่าตั้งต้น + ที่เพิ่มเอง + ที่เคยใช้ในข้อมูลเดิม (รวมจากทุกรายการในทุกใบ)
  // หมวดหมู่ใหญ่ทั้งหมด: มาจากฐานข้อมูล expenseCategories (key ของ object) — ไม่ต้องรวมจากประวัติอีกต่อไป เพราะฐานข้อมูลคือแหล่งความจริง
  const allMainCategories = Object.keys(expenseCategories || {});

  // หมวดหมู่ย่อยของหมวดหมู่ใหญ่ที่ระบุ: มาจากฐานข้อมูล expenseCategories โดยตรง
  const subCategoriesFor = (main) => ((expenseCategories || {})[main] || []).map((s) => (typeof s === "string" ? s : s.name));

  // เมื่อพิมพ์หมวดหมู่ใหญ่ใหม่ที่ยังไม่มี ให้เพิ่มเข้าฐานข้อมูล expenseCategories ทันที (ใช้ได้ทุกเครื่องหลังจากนี้)
  const handleItemMainCategoryChange = (idx, value) => {
    if (value && !allMainCategories.includes(value)) {
      setExpenseCategories((prev) => ({ ...prev, [value]: [] }));
    }
    updateItem(idx, "mainCategory", value); // updateItem จะรีเซ็ต subCategory ให้อัตโนมัติเมื่อเปลี่ยนหมวดหมู่ใหญ่
  };

  // เมื่อพิมพ์หมวดหมู่ย่อยใหม่ที่ยังไม่มีในหมวดหมู่ใหญ่นี้ ให้เพิ่มเข้าฐานข้อมูล expenseCategories ทันที
  const handleItemSubCategoryChange = (idx, mainCategory, value) => {
    if (value && mainCategory && !subCategoriesFor(mainCategory).includes(value)) {
      setExpenseCategories((prev) => ({ ...prev, [mainCategory]: [...(prev[mainCategory] || []), { name: value, openingBalance: 0, openingMonth: "" }] }));
    }
    updateItem(idx, "subCategory", value);
  };

  // โหมด "เพิ่มหมวดหมู่ใหม่" ต่อรายการ — เก็บเป็น { [idx]: "main" | "sub" | null }
  const [addingMainCatFor, setAddingMainCatFor] = useState({});
  const [addingSubCatFor, setAddingSubCatFor] = useState({});

  // --- คำนวณ VAT / หัก ณ ที่จ่าย / จำนวนเงินสุทธิ (คำนวณแยกต่อรายการ แล้วรวมผลลัพธ์) ---
  const calcTotals = (e) => {
    if (e.items && e.items.length > 0) {
      let amount = 0, vat = 0, wht = 0;
      e.items.forEach((it) => {
        const itAmount = Number(it.amount) || 0;
        amount += itAmount;
        vat += it.vatEnabled ? itAmount * 0.07 : 0;
        wht += itAmount * ((Number(it.whtRate) || 0) / 100);
      });
      return { amount, vat, wht, net: amount + vat - wht };
    }
    // รองรับข้อมูลเก่าที่ยังเป็นรายการเดียวระดับบิล
    const amount = Number(e.amount) || 0;
    const vat = e.vatEnabled ? amount * 0.07 : 0;
    const wht = amount * ((Number(e.whtRate) || 0) / 100);
    return { amount, vat, wht, net: amount + vat - wht };
  };

  const { amount: formAmount, vat: formVat, wht: formWht, net: formNet } = calcTotals(form);
  const formPaid = (form.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const formRemaining = formNet - formPaid;

  const addItem = () => setForm({ ...form, items: [...(form.items || []), blankItem()] });
  const updateItem = (idx, field, value) => {
    const items = [...(form.items || [])];
    let updated = { ...items[idx], [field]: value };
    // ถ้าเปลี่ยนหมวดหมู่ใหญ่ ให้รีเซ็ตหมวดหมู่ย่อยเป็นตัวเลือกแรกของหมวดใหม่
    if (field === "mainCategory") {
      updated.subCategory = "";
    }
    items[idx] = updated;
    setForm({ ...form, items });
  };
  const removeItem = (idx) => setForm({ ...form, items: (form.items || []).filter((_, i) => i !== idx) });

  const save = () => {
    const items = (form.items || []).map((it) => ({ ...it, amount: Number(it.amount) || 0, vatEnabled: !!it.vatEnabled, whtRate: Number(it.whtRate) || 0 }));
    const totalAmount = items.reduce((s, it) => s + it.amount, 0);
    if (!(totalAmount > 0)) return;
    const cleaned = {
      ...form,
      items,
      amount: totalAmount, // เก็บยอดรวมไว้ที่ระดับบนสุดด้วย เพื่อความเข้ากันได้กับการคำนวณเดิม (dashboard ฯลฯ)
      payments: (form.payments || []).map((p) => ({ ...p, amount: Number(p.amount) || 0 })),
      loanInstallment: pendingInstallment || form.loanInstallment || null,
    };
    if (modal.mode === "add") setExpenses([...expenses, cleaned]);
    else setExpenses(expenses.map((e) => (e.id === modal.item.id ? cleaned : e)));

    // ถ้าผูกกับงวดผ่อน ให้บันทึกสถานะ "จ่ายแล้ว" ในสัญญาเงินกู้ด้วย
    if (cleaned.loanInstallment && setLoans) {
      setLoans((loans || []).map((loan) => {
        if (loan.id !== cleaned.loanInstallment.loanId) return loan;
        const already = (loan.paidInstallments || []).some((p) => p.no === cleaned.loanInstallment.no);
        if (already) return loan;
        return {
          ...loan,
          paidInstallments: [...(loan.paidInstallments || []), { no: cleaned.loanInstallment.no, expenseId: cleaned.id, paidDate: cleaned.billDate || cleaned.recordDate }],
        };
      }));
    }

    setPendingInstallment(null);
    setModal(null);
  };

  const remove = (id) => setExpenses(expenses.filter((e) => e.id !== id));

  const fromLabel = (id) => {
    if (id === "CASH") return "เงินสดหน้าร้าน";
    const b = storeBankAccounts.find((b) => b.id === id);
    return b ? `${b.bankName} ${b.accountNo}` : "-";
  };

  const filtered = expenses
    .filter((e) => {
      const items = (e.items && e.items.length > 0) ? e.items : [{ description: e.description, mainCategory: e.mainCategory || e.category, subCategory: e.subCategory }];
      const itemMatch = items.some((it) => (it.mainCategory || "").includes(search) || (it.subCategory || "").includes(search) || (it.description || "").includes(search));
      return itemMatch || e.id.includes(search) || (e.refNo || "").includes(search) || (e.taxInvoiceNo || "").includes(search);
    })
    .filter((e) => (!dateFrom || (e.billDate || e.date || "") >= dateFrom) && (!dateTo || (e.billDate || e.date || "") <= dateTo))
    .sort((a, b) => ((a.billDate || a.date) < (b.billDate || b.date) ? 1 : (a.billDate || a.date) > (b.billDate || b.date) ? -1 : 0));

  const totalAll = expenses.reduce((s, e) => s + calcTotals(e).net, 0);

  // เดือนปัจจุบัน (ตามรูปแบบ YYYY-MM ของวันที่ตามบิล)
  const currentMonth = new Date().toISOString().slice(0, 7);
  const totalThisMonth = expenses.filter((e) => ((e.billDate || e.date) || "").startsWith(currentMonth)).reduce((s, e) => s + calcTotals(e).net, 0);

  // สรุปตามหมวดหมู่ใหญ่ พร้อม breakdown หมวดหมู่ย่อย
  const byCategory = useMemo(() => {
    const groups = {}; // mainCategory -> { total, subs: { subCategory: total } }
    expenses.forEach((e) => {
      const t = calcTotals(e);
      const items = (e.items && e.items.length > 0) ? e.items : [{ description: e.description, mainCategory: e.mainCategory || e.category, subCategory: e.subCategory, amount: e.amount }];
      items.forEach((it) => {
        const main = it.mainCategory || "ไม่ระบุ";
        const sub = it.subCategory || "ไม่ระบุ";
        // กระจายยอดสุทธิ (รวม VAT/หัก ณ ที่จ่าย) ของใบนี้ตามสัดส่วนจำนวนเงินของแต่ละรายการ
        const itemAmount = Number(it.amount) || 0;
        const share = t.amount > 0 ? itemAmount / t.amount : 0;
        const net = t.net * share;
        if (!groups[main]) groups[main] = { total: 0, subs: {} };
        groups[main].total += net;
        groups[main].subs[sub] = (groups[main].subs[sub] || 0) + net;
      });
    });
    return Object.entries(groups)
      .map(([mainCategory, g]) => ({
        mainCategory,
        total: g.total,
        subs: Object.entries(g.subs).map(([subCategory, amount]) => ({ subCategory, amount })).sort((a, b) => b.amount - a.amount),
      }))
      .sort((a, b) => b.total - a.total);
  }, [expenses]);

  return (
    <div>
      <Header title="บันทึกค่าใช้จ่าย" subtitle="บันทึกค่าใช้จ่ายในการดำเนินงานของร้าน พร้อมพิมพ์ใบสำคัญจ่าย">
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <ExportToolbar
            onPDF={() => printAsPDF("tab-export-expenses", "ค่าใช้จ่าย")}
            onExcel={() => {
              const rows = [
                ["เลขที่อ้างอิง", "วันที่ตามบิล", "เลขที่ใบกำกับ", "หมวดหมู่ใหญ่", "หมวดหมู่ย่อย", "รายละเอียด", "จำนวนเงิน", "VAT", "หัก ณ ที่จ่าย", "สุทธิ"],
                ...filtered.map((e) => {
                  const items = (e.items && e.items.length > 0) ? e.items : [{ mainCategory: e.mainCategory || e.category, subCategory: e.subCategory, description: e.description, amount: e.amount, vatEnabled: e.vatEnabled, whtRate: e.whtRate }];
                  const { amount, vat, wht } = calcTotals({ ...e, items });
                  return [e.refNo || e.id, e.billDate || e.date, e.taxInvoiceNo || "", items.map(it=>it.mainCategory).join(", "), items.map(it=>it.subCategory).join(", "), items.map(it=>it.description).join(", "), amount, vat, wht, amount + vat - wht];
                }),
              ];
              exportExcel(rows, "ค่าใช้จ่าย.xlsx", "ค่าใช้จ่าย");
            }}
            onImage={() => printAsPDF("tab-export-expenses", "ค่าใช้จ่าย")}
          />
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> บันทึกค่าใช้จ่าย</button>
        </div>
      </Header>
      <div id="tab-export-expenses">
      {/* การ์ดสรุป */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: 12, marginBottom: 20 }}>
        {/* การ์ดเดือนนี้ */}
        <div style={{ background: "#faece7", borderRadius: 12, border: "1px solid #f0c0a0", padding: "14px 16px" }}>
          <div style={{ fontSize: 11, color: "#993c1d", marginBottom: 4, fontWeight: 600 }}>เดือนนี้ (สุทธิ)</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: "#993c1d" }}>฿{fmt(totalThisMonth)}</div>
        </div>
        {/* การ์ดรวมทั้งหมด */}
        <div style={{ background: "#f1efe8", borderRadius: 12, border: "1px solid #d4d0c0", padding: "14px 16px" }}>
          <div style={{ fontSize: 11, color: "#444441", marginBottom: 4, fontWeight: 600 }}>รวมทั้งหมด (สุทธิ)</div>
          <div style={{ fontSize: 20, fontWeight: 700, color: "#444441" }}>฿{fmt(totalAll)}</div>
        </div>
        {/* การ์ดแต่ละหมวดหมู่ใหญ่ */}
        {byCategory.map((c) => (
          <div key={c.mainCategory} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "14px 16px" }}>
            <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 4, fontWeight: 600 }}>{c.mainCategory}</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: "#374151" }}>฿{fmt(c.total)}</div>
          </div>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr", gap: 16, alignItems: "start" }}>
        <div>
          <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาเลขที่ใบกำกับภาษี, หมวดหมู่ หรือรายละเอียด..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />
          <Card>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={thStyle}>เลขที่อ้างอิง</th>
                  <th style={thStyle}>วันที่ตามบิล</th>
                  <th style={thStyle}>เลขที่ใบกำกับภาษี</th>
                  <th style={thStyle}>ผู้รับเงิน</th>
                  <th style={thStyle}>หมวดหมู่ / รายละเอียด</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงินสุทธิ</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((e) => {
                  const t = calcTotals(e);
                  const items = (e.items && e.items.length > 0) ? e.items : [{ description: e.description, mainCategory: e.mainCategory || e.category, subCategory: e.subCategory, amount: e.amount }];
                  // จัดกลุ่มตามหมวดหมู่ย่อย
                  const isExpanded = !!expandedExpenses[e.id];
                  return (
                    <React.Fragment key={e.id}>
                      <tr
                        onClick={() => toggleExpense(e.id)}
                        style={{ cursor: "pointer" }}
                        onMouseEnter={(e2) => e2.currentTarget.style.background = "#f9fafb"}
                        onMouseLeave={(e2) => e2.currentTarget.style.background = ""}
                      >
                        <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
                            {isExpanded ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                            {e.refNo || e.id}
                          </div>
                        </td>
                        <td style={tdStyle}>{e.billDate || e.date}</td>
                        <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace" }}>{e.taxInvoiceNo || "-"}</td>
                        <td style={{ ...tdStyle, fontWeight: 600, color: "#374151" }}>{e.vendorName || "-"}</td>
                        <td style={tdStyle}></td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#993c1d" }}>-฿{fmt(t.net)}</td>
                        <td style={{ ...tdStyle, textAlign: "right" }} onClick={(e2) => e2.stopPropagation()}>
                          <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                            <button style={iconBtn} onClick={() => openView(e)}><Printer size={14} /> ใบสำคัญจ่าย</button>
                            <button style={iconBtn} onClick={() => openEdit(e)}><Edit2 size={14} /> แก้ไข</button>
                            <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบรายการค่าใช้จ่าย "${e.refNo || e.id}" ใช่หรือไม่?`, () => remove(e.id))}><Trash2 size={14} /> ลบ</button>
                          </div>
                        </td>
                      </tr>
                      {isExpanded && items.map((it, ii) => (
                        <tr key={ii} style={{ background: "#f9fafb" }}>
                          <td style={{ ...tdStyle, paddingLeft: 28, color: "#9ca3af" }}>↳</td>
                          <td style={tdStyle}></td>
                          <td style={tdStyle}></td>
                          <td style={tdStyle}></td>
                          <td style={{ ...tdStyle, color: "#6b7280" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: it.description ? 2 : 0 }}>
                              <Badge text={it.mainCategory || "-"} />
                              {it.subCategory && <span style={{ fontSize: 12, color: "#9ca3af" }}>{it.subCategory}</span>}
                            </div>
                            {it.description && <div style={{ fontSize: 13, paddingLeft: 2 }}>{it.description}</div>}
                          </td>
                          <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d" }}>-฿{fmt(it.amount)}</td>
                          <td style={tdStyle}></td>
                        </tr>
                      ))}
                    </React.Fragment>
                  );
                })}
                {filtered.length === 0 && <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีรายการค่าใช้จ่าย</td></tr>}
              </tbody>
            </table>
          </Card>
        </div>

        <div>
          <h3 style={{ fontSize: 15, fontWeight: 600, margin: "0 0 10px" }}>สรุปตามหมวดหมู่ (สุทธิ)</h3>
          <Card>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={thStyle}>หมวดหมู่</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>รวม</th>
                </tr>
              </thead>
              <tbody>
                {byCategory.map((c) => (
                  <React.Fragment key={c.mainCategory}>
                    <tr>
                      <td style={{ ...tdStyle, fontWeight: 700 }}>{c.mainCategory}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>฿{fmt(c.total)}</td>
                    </tr>
                    {c.subs.map((s) => (
                      <tr key={c.mainCategory + "__" + s.subCategory}>
                        <td style={{ ...tdStyle, paddingLeft: 28, color: "#6b7280", fontSize: 13 }}>{s.subCategory}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280", fontSize: 13 }}>฿{fmt(s.amount)}</td>
                      </tr>
                    ))}
                  </React.Fragment>
                ))}
                {byCategory.length === 0 && <tr><td colSpan={2} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีข้อมูล</td></tr>}
              </tbody>
            </table>
          </Card>
        </div>
      </div>

      {modal && (modal.mode === "add" || modal.mode === "edit") && (
        <Modal title={modal.mode === "add" ? "บันทึกค่าใช้จ่าย" : `แก้ไขค่าใช้จ่าย · ${form.refNo || form.id}`} onClose={() => setModal(null)} wide fullscreen>
          {unpaidInstallments.length > 0 && (
            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 8 }}>
              <button style={btnSecondary} onClick={() => { setPickerLoanId(""); setInstallmentPicker(true); }}><CreditCard size={14} /> ดึงจากงวดผ่อน</button>
            </div>
          )}
          {pendingInstallment && (
            <div style={{ background: "#eeedfe", borderRadius: 8, padding: "8px 12px", marginBottom: 12, fontSize: 13, color: "#3c3489", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <span>เชื่อมกับงวดผ่อนสัญญา: {(loans || []).find((l) => l.id === pendingInstallment.loanId)?.name || pendingInstallment.loanId} · งวดที่ {pendingInstallment.no}</span>
              <button style={{ ...btnSecondary, padding: "4px 8px" }} onClick={() => setPendingInstallment(null)}>ยกเลิกการเชื่อม</button>
            </div>
          )}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1.5fr", gap: "0 16px" }}>
            <Field label="เลขที่อ้างอิง">
              <input style={inputStyle} value={form.refNo} onChange={(e) => setForm({ ...form, refNo: e.target.value })} placeholder="รันอัตโนมัติ — แก้ไขได้" />
            </Field>
            <Field label="วันที่บันทึก">
              <input type="date" style={inputStyle} value={form.recordDate} onChange={(e) => setForm({ ...form, recordDate: e.target.value })} />
            </Field>
            <Field label="วันที่ตามบิล">
              <input type="date" style={inputStyle} value={form.billDate} onChange={(e) => {
                const newDate = e.target.value;
                const newRefNo = genId("EX", expenses, newDate);
                setForm({ ...form, billDate: newDate, refNo: newRefNo });
              }} />
            </Field>
            <Field label="ผู้รับเงิน / ร้านค้า">
              <CustomerSelect
                customers={customers || []}
                value={form.vendorId || ""}
                onChange={(cid) => {
                  const c = (customers || []).find((x) => x.id === cid);
                  setForm({ ...form, vendorId: cid, vendorName: c ? c.name : "" });
                }}
                labelWithId={false}
              />
            </Field>
          </div>

          <Field label="เลขที่ใบกำกับภาษี">
            <input style={inputStyle} value={form.taxInvoiceNo} onChange={(e) => setForm({ ...form, taxInvoiceNo: e.target.value })} placeholder="เช่น INV-1234" />
          </Field>

          <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 8 }}>รายการค่าใช้จ่าย (เพิ่มได้หลายรายการในใบเดียว)</div>
          {(form.items || []).map((it, idx) => (
            <div key={it.id} style={{ border: "1px solid #e5e7eb", borderRadius: 8, padding: 12, marginBottom: 10, background: "#f9fafb" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontWeight: 600, fontSize: 13, color: "#6b7280" }}>รายการที่ {idx + 1}</div>
                {(form.items || []).length > 1 && (
                  <button style={btnDanger} onClick={() => removeItem(idx)}><Trash2 size={14} /> ลบรายการ</button>
                )}
              </div>
              {/* แถว 1: หมวดใหญ่ หมวดย่อย รายละเอียด */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 2fr", gap: "0 12px", marginBottom: 8 }}>
                <Field label="หมวดหมู่ใหญ่">
                  {addingMainCatFor[idx] ? (
                    <div style={{ display: "flex", gap: 6 }}>
                      <input style={inputStyle} autoFocus value={it.mainCategory} onChange={(e) => handleItemMainCategoryChange(idx, e.target.value)} placeholder="พิมพ์ชื่อหมวดหมู่ใหม่" />
                      <button type="button" style={roundBtn} title="เสร็จสิ้น" onClick={() => setAddingMainCatFor((p) => ({ ...p, [idx]: false }))}><Check size={14} /></button>
                    </div>
                  ) : (
                    <select style={inputStyle} value={it.mainCategory} onChange={(e) => {
                      if (e.target.value === "__add_new__") { setAddingMainCatFor((p) => ({ ...p, [idx]: true })); handleItemMainCategoryChange(idx, ""); return; }
                      handleItemMainCategoryChange(idx, e.target.value);
                    }}>
                      <option value="">-- เลือกหมวดหมู่ใหญ่ --</option>
                      {allMainCategories.map((c) => <option key={c} value={c}>{c}</option>)}
                      <option value="__add_new__">+ เพิ่มหมวดหมู่ใหม่...</option>
                    </select>
                  )}
                </Field>
                <Field label="หมวดหมู่ย่อย">
                  {addingSubCatFor[idx] ? (
                    <div style={{ display: "flex", gap: 6 }}>
                      <input style={inputStyle} autoFocus value={it.subCategory} onChange={(e) => handleItemSubCategoryChange(idx, it.mainCategory, e.target.value)} placeholder="พิมพ์ชื่อหมวดหมู่ย่อยใหม่" />
                      <button type="button" style={roundBtn} title="เสร็จสิ้น" onClick={() => setAddingSubCatFor((p) => ({ ...p, [idx]: false }))}><Check size={14} /></button>
                    </div>
                  ) : (
                    <select style={inputStyle} value={it.subCategory} onChange={(e) => {
                      if (e.target.value === "__add_new__") { setAddingSubCatFor((p) => ({ ...p, [idx]: true })); handleItemSubCategoryChange(idx, it.mainCategory, ""); return; }
                      handleItemSubCategoryChange(idx, it.mainCategory, e.target.value);
                    }} disabled={!it.mainCategory}>
                      <option value="">-- เลือกหมวดหมู่ย่อย --</option>
                      {subCategoriesFor(it.mainCategory).map((c) => <option key={c} value={c}>{c}</option>)}
                      <option value="__add_new__">+ เพิ่มหมวดหมู่ย่อยใหม่...</option>
                    </select>
                  )}
                </Field>
                <Field label="รายละเอียด">
                  <input style={inputStyle} value={it.description} onChange={(e) => updateItem(idx, "description", e.target.value)} placeholder="เช่น ค่าน้ำมันรถบรรทุกขนของ" />
                </Field>
              </div>
              {/* แถว 2: จำนวนเงิน VAT ✓ จำนวนVAT หัก ณ ที่จ่าย จำนวนหัก ณ ที่จ่าย สุทธิ */}
              {(() => {
                const amt = Number(it.amount) || 0;
                const vatAmt = it.vatEnabled ? amt * 0.07 : 0;
                const whtAmt = amt * ((Number(it.whtRate) || 0) / 100);
                const net = amt + vatAmt - whtAmt;
                return (
                  <div style={{ display: "grid", gridTemplateColumns: "1.2fr 0.7fr 0.8fr 0.8fr 0.8fr 0.8fr", gap: "0 12px", alignItems: "end" }}>
                    <Field label="จำนวนเงิน (บาท)">
                      <input type="number" style={{ ...inputStyle, textAlign: "right" }} value={it.amount} onChange={(e) => updateItem(idx, "amount", e.target.value)} placeholder="0" />
                    </Field>
                    <Field label="VAT 7%">
                      <label style={{ display: "flex", alignItems: "center", gap: 6, height: 38, fontSize: 14, cursor: "pointer" }}>
                        <input type="checkbox" checked={!!it.vatEnabled} onChange={(e) => updateItem(idx, "vatEnabled", e.target.checked)} style={{ width: 16, height: 16 }} />
                        ติ๊ก
                      </label>
                    </Field>
                    <Field label="จำนวน VAT">
                      <div style={{ ...inputStyle, background: "#f3f4f6", textAlign: "right", color: vatAmt > 0 ? "#185fa5" : "#9ca3af" }}>
                        {vatAmt > 0 ? `+${fmt(vatAmt)}` : "0.00"}
                      </div>
                    </Field>
                    <Field label="หัก ณ ที่จ่าย (%)">
                      <input type="number" style={{ ...inputStyle, textAlign: "right" }} value={it.whtRate} onChange={(e) => updateItem(idx, "whtRate", e.target.value)} placeholder="0" />
                    </Field>
                    <Field label="จำนวนหัก ณ ที่จ่าย">
                      <div style={{ ...inputStyle, background: "#f3f4f6", textAlign: "right", color: whtAmt > 0 ? "#993c1d" : "#9ca3af" }}>
                        {whtAmt > 0 ? `-${fmt(whtAmt)}` : "0.00"}
                      </div>
                    </Field>
                    <Field label="จำนวนเงินสุทธิ">
                      <div style={{ ...inputStyle, background: "#e1f5ee", textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>
                        {fmt(net)}
                      </div>
                    </Field>
                  </div>
                );
              })()}
            </div>
          ))}
          <div style={{ marginBottom: 10 }}>
            <button style={btnSecondary} onClick={addItem}><Plus size={14} /> เพิ่มรายการ</button>
          </div>
          <p style={{ fontSize: 12, color: "#9ca3af", marginTop: -2, marginBottom: 16 }}>* เลือกหมวดหมู่จากรายการ หรือกด "+ เพิ่มหมวดหมู่ใหม่..." เพื่อพิมพ์ชื่อใหม่ ระบบจะบันทึกเป็นหมวดหมู่ใหม่ในฐานข้อมูลให้อัตโนมัติ — แต่ละรายการตั้งค่า VAT/หัก ณ ที่จ่ายแยกกันได้</p>

          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", marginBottom: 16, fontSize: 14 }}>
            <Row label="จำนวนเงินรวมทุกรายการ" value={`฿${fmt(formAmount)}`} />
            <Row label="ภาษีมูลค่าเพิ่มรวม" value={`+฿${fmt(formVat)}`} />
            <Row label="หัก ณ ที่จ่ายรวม" value={`-฿${fmt(formWht)}`} />
            <Row label="จำนวนเงินสุทธิ" value={`฿${fmt(formNet)}`} bold />
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8, marginBottom: 8 }}>
            <div style={{ fontWeight: 600, fontSize: 14 }}>การชำระเงิน</div>
          </div>

          <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "12px 16px", fontSize: 13, color: "#854f0b" }}>
            บันทึกการจ่ายเงินสำหรับบิลนี้ได้ที่หน้า <strong>"รับชำระ / จ่ายชำระ"</strong> แทน — กดปุ่ม "ค้างจ่าย (ค่าใช้จ่าย)" เพื่อหารายการนี้และบันทึกจ่ายได้ที่นั่น
          </div>

          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "10px 16px", marginTop: 12, fontSize: 13 }}>
            <Row label="ยอดที่ต้องชำระ" value={`฿${fmt(formNet)}`} />
            <Row label="ชำระแล้ว" value={`฿${fmt(formPaid)}`} />
            <Row label="คงค้าง" value={`฿${fmt(formRemaining)}`} bold color={formRemaining > 0 ? "#a32d2d" : "#27500a"} />
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {modal && modal.mode === "view" && (
        <ExpenseVoucherModal expense={modal.item} storeBankAccounts={storeBankAccounts} companySettings={companySettings} onClose={() => setModal(null)} />
      )}

      {installmentPicker && (
        <Modal title="ดึงจากงวดผ่อน — เลือกงวดที่ยังไม่ชำระ" onClose={() => setInstallmentPicker(false)} wide>
          {unpaidInstallments.length === 0 ? (
            <p style={{ color: "#9ca3af", textAlign: "center", padding: "24px 0" }}>ไม่มีงวดผ่อนที่ยังไม่ชำระ</p>
          ) : !pickerLoanId ? (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr>
                    <th style={thStyle}>เลขที่สัญญา</th>
                    <th style={thStyle}>ชื่อสัญญา</th>
                    <th style={thStyle}>ประเภท</th>
                    <th style={thStyle}>วันครบกำหนดชำระ</th>
                    <th style={thStyle}></th>
                  </tr>
                </thead>
                <tbody>
                  {loansWithUnpaid.map((loan) => (
                    <tr
                      key={loan.id}
                      style={{ cursor: "pointer" }}
                      onClick={() => setPickerLoanId(loan.id)}
                      onMouseEnter={(e) => { e.currentTarget.style.background = "#f3f4f6"; }}
                      onMouseLeave={(e) => { e.currentTarget.style.background = ""; }}
                    >
                      <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{loan.billNo || loan.id}</td>
                      <td style={tdStyle}>{loan.name}</td>
                      <td style={tdStyle}><Badge text={loan.type} /></td>
                      <td style={tdStyle}>{loan.nextDueDate}</td>
                      <td style={{ ...tdStyle, textAlign: "right", color: "#534ab7" }}><ChevronRight size={16} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <>
              {(() => {
                const currentLoan = loansWithUnpaid.find((l) => l.id === pickerLoanId);
                return (
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                    <button style={btnSecondary} onClick={() => setPickerLoanId("")}>‹ กลับไปเลือกสัญญา</button>
                    <div style={{ fontWeight: 600 }}>
                      {currentLoan?.name}{currentLoan?.billNo ? ` · ${currentLoan.billNo}` : ""} <Badge text={currentLoan?.type || ""} />
                    </div>
                  </div>
                );
              })()}
              <div style={{ overflowX: "auto", maxHeight: 360, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr>
                      <th style={thStyle}>งวดที่</th>
                      <th style={thStyle}>วันครบกำหนด</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>ยอดผ่อน</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>ดอกเบี้ย</th>
                      <th style={{ ...thStyle, textAlign: "right" }}>ตัดเงินต้น</th>
                      <th style={thStyle}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {pickerInstallments.map(({ loan, installment }) => (
                      <tr key={loan.id + "-" + installment.no}>
                        <td style={tdStyle}>{installment.no} / {loan.totalInstallments}</td>
                        <td style={tdStyle}>{installment.dueDate}</td>
                        <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>฿{fmt(installment.payment)}</td>
                        <td style={{ ...tdStyle, textAlign: "right", color: "#854f0b" }}>฿{fmt(installment.interest)}</td>
                        <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(installment.principalPortion)}</td>
                        <td style={tdStyle}>
                          <button style={btnPrimary} onClick={() => applyInstallment(loan, installment)}>เลือก</button>
                        </td>
                      </tr>
                    ))}
                    {pickerInstallments.length === 0 && (
                      <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีงวดค้างชำระสำหรับสัญญานี้</td></tr>
                    )}
                  </tbody>
                </table>
              </div>
            </>
          )}
          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setInstallmentPicker(false)}>ปิด</button>
          </div>
        </Modal>
      )}
      </div>{/* end tab-export-expenses */}
    </div>
  );
}
// ใบสำคัญจ่าย (Payment Voucher) PDF view
function ExpenseVoucherModal({ expense, storeBankAccounts, companySettings, onClose }) {
  const cs = companySettings || {};
  const items = (expense.items && expense.items.length > 0)
    ? expense.items
    : [{ description: expense.description, mainCategory: expense.mainCategory || expense.category, subCategory: expense.subCategory, amount: expense.amount, vatEnabled: expense.vatEnabled, whtRate: expense.whtRate }];
  const amount = items.reduce((s, it) => s + (Number(it.amount) || 0), 0);
  const vat = items.reduce((s, it) => s + (it.vatEnabled ? (Number(it.amount) || 0) * 0.07 : 0), 0);
  const wht = items.reduce((s, it) => s + (Number(it.amount) || 0) * ((Number(it.whtRate) || 0) / 100), 0);
  const net = amount + vat - wht;
  const payments = expense.payments || [];
  const totalPaid = payments.reduce((s, p) => s + (Number(p.amount) || 0), 0);

  const fromLabel = (id) => {
    if (id === "CASH") return "เงินสดหน้าร้าน";
    const b = (storeBankAccounts || []).find((b) => b.id === id);
    return b ? `${b.bankName} ${b.accountNo}` : "-";
  };

  return (
    <Modal title={`ใบสำคัญจ่าย · ${expense.refNo || expense.id}`} onClose={onClose} wide>
      <div id="expense-voucher-pdf-content" style={{ background: "#fff", padding: "24px", border: "1px solid #e5e7eb", borderRadius: 8 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${cs.accentColor || "#993c1d"}`, paddingBottom: 12, marginBottom: 16 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {cs.logo && (
              <img src={cs.logo} alt="logo" style={{ height: 50, maxWidth: 100, objectFit: "contain" }} />
            )}
            <div>
              <div style={{ fontSize: 18, fontWeight: 700, color: cs.accentColor || "#993c1d" }}>{cs.name || "วงจรกรีน รีไซเคิล"}</div>
              {cs.taxId && <div style={{ fontSize: 12, color: "#6b7280" }}>เลขผู้เสียภาษี: {cs.taxId}</div>}
              {cs.address && <div style={{ fontSize: 12, color: "#6b7280" }}>{cs.address}</div>}
              {cs.phone && <div style={{ fontSize: 12, color: "#6b7280" }}>โทร: {cs.phone}</div>}
            </div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 16, fontWeight: 700 }}>ใบสำคัญจ่าย</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>เลขที่อ้างอิง: {expense.refNo || expense.id}</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>วันที่บันทึก: {expense.recordDate || expense.billDate || expense.date}</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>วันที่ตามบิล: {expense.billDate || expense.date}</div>
            {expense.taxInvoiceNo && <div style={{ fontSize: 12, color: "#6b7280" }}>เลขที่ใบกำกับภาษี: {expense.taxInvoiceNo}</div>}
          </div>
        </div>

        <div style={{ marginBottom: 16, fontSize: 13 }}>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>รายละเอียดค่าใช้จ่าย</div>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr style={{ background: "#f3f4f6" }}>
                <th style={thStyle}>รายละเอียด</th>
                <th style={thStyle}>หมวดหมู่</th>
                <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
                <th style={{ ...thStyle, textAlign: "center" }}>VAT</th>
                <th style={{ ...thStyle, textAlign: "center" }}>หัก ณ ที่จ่าย</th>
              </tr>
            </thead>
            <tbody>
              {items.map((it, i) => (
                <tr key={i}>
                  <td style={tdStyle}>{it.description || "-"}</td>
                  <td style={tdStyle}>
                    <Badge text={it.mainCategory || "-"} />
                    {it.subCategory && <span style={{ marginLeft: 6, color: "#6b7280" }}>› {it.subCategory}</span>}
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(it.amount)}</td>
                  <td style={{ ...tdStyle, textAlign: "center" }}>{it.vatEnabled ? "7%" : "-"}</td>
                  <td style={{ ...tdStyle, textAlign: "center" }}>{Number(it.whtRate) > 0 ? `${it.whtRate}%` : "-"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end" }}>
          <div style={{ width: 280 }}>
            <Row label="จำนวนเงิน" value={`฿${fmt(amount)}`} />
            <Row label="ภาษีมูลค่าเพิ่มรวม" value={`+฿${fmt(vat)}`} />
            <Row label="หัก ณ ที่จ่ายรวม" value={`-฿${fmt(wht)}`} />
            <div style={{ display: "flex", justifyContent: "space-between", padding: "6px 0", borderTop: "2px solid #993c1d", fontWeight: 700, fontSize: 15 }}>
              <span>จำนวนเงินสุทธิ</span>
              <span>฿{fmt(net)}</span>
            </div>
          </div>
        </div>

        {payments.length > 0 && (
          <div style={{ marginTop: 16 }}>
            <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 6 }}>รายการจ่ายชำระเงิน</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: "#f3f4f6" }}>
                  <th style={thStyle}>วันที่จ่าย</th>
                  <th style={thStyle}>จ่ายจาก</th>
                  <th style={thStyle}>วิธีชำระ</th>
                  <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงิน</th>
                </tr>
              </thead>
              <tbody>
                {payments.map((p) => (
                  <tr key={p.id}>
                    <td style={tdStyle}>{p.date}</td>
                    <td style={tdStyle}>{fromLabel(p.fromStoreBankId)}</td>
                    <td style={tdStyle}>{p.method}</td>
                    <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(p.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{ textAlign: "right", fontSize: 12, marginTop: 6, fontWeight: 600 }}>
              ชำระแล้วทั้งหมด: ฿{fmt(totalPaid)} / คงเหลือ: ฿{fmt(net - totalPaid)}
            </div>
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "space-between", marginTop: 48, fontSize: 12 }}>
          <div style={{ textAlign: "center", width: "30%" }}>
            <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้จัดทำ</div>
          </div>
          <div style={{ textAlign: "center", width: "30%" }}>
            <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้อนุมัติ</div>
          </div>
          <div style={{ textAlign: "center", width: "30%" }}>
            <div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้รับเงิน</div>
          </div>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
        <button style={btnSecondary} onClick={onClose}>ปิด</button>
        <button style={btnPrimary} onClick={() => printAsPDF("expense-voucher-pdf-content", `ใบสำคัญจ่าย ${expense.refNo || expense.id}`)}><Download size={16} /> พิมพ์ / บันทึก PDF</button>
      </div>
    </Modal>
  );
}

// ===================================================================
// EXPENSE CATEGORIES TAB (หมวดหมู่ค่าใช้จ่าย — ฐานข้อมูลแยกต่างหาก)
// ===================================================================
function ExpenseCategoriesTab({ expenseCategories, setExpenseCategories, expenses, setExpenses }) {
  const [search, setSearch] = useState("");
  const [mainModal, setMainModal] = useState(null); // {mode:'add'|'edit', oldName}
  const [mainName, setMainName] = useState("");
  const [subModal, setSubModal] = useState(null); // {mode:'add'|'edit', main, oldName}
  const [subForm, setSubForm] = useState({ name: "", openingBalance: 0, openingMonth: "" });

  const mains = Object.keys(expenseCategories || {});

  // รองรับข้อมูลเดิมที่หมวดหมู่ย่อยยังเป็น string ธรรมดา (ก่อนมีฟีเจอร์ยอดยกมา) — แปลงเป็น object ให้อัตโนมัติตอนอ่าน
  const subsOf = (main) => (expenseCategories[main] || []).map((s) => (typeof s === "string" ? { name: s, openingBalance: 0, openingMonth: "" } : s));

  const itemsOf = (e) => (e.items && e.items.length > 0) ? e.items : [{ mainCategory: e.mainCategory || e.category, subCategory: e.subCategory }];

  const countMain = (main) => expenses.filter((e) => itemsOf(e).some((it) => it.mainCategory === main)).length;
  const countSub = (main, sub) => expenses.filter((e) => itemsOf(e).some((it) => it.mainCategory === main && it.subCategory === sub)).length;

  const filtered = mains.filter((m) => m.includes(search) || subsOf(m).some((s) => s.name.includes(search)));

  const monthLabelOf = (ym) => {
    if (!ym) return "";
    const [y, m] = ym.split("-");
    return `${MONTH_NAMES_TH[Number(m)]} ${y}`;
  };

  // ---------- หมวดหมู่ใหญ่ ----------
  const openAddMain = () => { setMainName(""); setMainModal({ mode: "add" }); };
  const openEditMain = (main) => { setMainName(main); setMainModal({ mode: "edit", oldName: main }); };

  const saveMain = () => {
    const trimmed = mainName.trim();
    if (!trimmed) return;
    if (mainModal.mode === "add") {
      if (expenseCategories[trimmed]) { alert("มีหมวดหมู่ใหญ่นี้อยู่แล้ว"); return; }
      setExpenseCategories({ ...expenseCategories, [trimmed]: [] });
    } else {
      if (trimmed !== mainModal.oldName && expenseCategories[trimmed]) { alert("มีหมวดหมู่ใหญ่นี้อยู่แล้ว"); return; }
      const updated = { ...expenseCategories };
      const subs = updated[mainModal.oldName] || [];
      delete updated[mainModal.oldName];
      updated[trimmed] = subs;
      setExpenseCategories(updated);
      if (trimmed !== mainModal.oldName) {
        setExpenses(expenses.map((e) => {
          if (!(e.items && e.items.length > 0)) {
            return (e.mainCategory === mainModal.oldName) ? { ...e, mainCategory: trimmed } : e;
          }
          return { ...e, items: e.items.map((it) => it.mainCategory === mainModal.oldName ? { ...it, mainCategory: trimmed } : it) };
        }));
      }
    }
    setMainModal(null);
  };

  const removeMain = (main) => {
    const used = countMain(main);
    if (used > 0) {
      alert(`ลบไม่ได้ — มีค่าใช้จ่าย ${used} รายการที่ใช้หมวดหมู่นี้อยู่ กรุณาเปลี่ยนหมวดหมู่ของรายการนั้นก่อน`);
      return;
    }
    const updated = { ...expenseCategories };
    delete updated[main];
    setExpenseCategories(updated);
  };

  // ---------- หมวดหมู่ย่อย ----------
  const openAddSub = (main) => { setSubForm({ name: "", openingBalance: 0, openingMonth: "" }); setSubModal({ mode: "add", main }); };
  const openEditSub = (main, sub) => { setSubForm({ name: sub.name, openingBalance: sub.openingBalance || 0, openingMonth: sub.openingMonth || "" }); setSubModal({ mode: "edit", main, oldName: sub.name }); };

  const saveSub = () => {
    const trimmed = subForm.name.trim();
    if (!trimmed || !subModal) return;
    const main = subModal.main;
    const subs = subsOf(main);
    const cleaned = { name: trimmed, openingBalance: Number(subForm.openingBalance) || 0, openingMonth: subForm.openingMonth || "" };
    if (subModal.mode === "add") {
      if (subs.some((s) => s.name === trimmed)) { alert("มีหมวดหมู่ย่อยนี้อยู่แล้ว"); return; }
      setExpenseCategories({ ...expenseCategories, [main]: [...subs, cleaned] });
    } else {
      if (trimmed !== subModal.oldName && subs.some((s) => s.name === trimmed)) { alert("มีหมวดหมู่ย่อยนี้อยู่แล้ว"); return; }
      setExpenseCategories({ ...expenseCategories, [main]: subs.map((s) => (s.name === subModal.oldName ? cleaned : s)) });
      if (trimmed !== subModal.oldName) {
        setExpenses(expenses.map((e) => {
          if (!(e.items && e.items.length > 0)) {
            return (e.mainCategory === main && e.subCategory === subModal.oldName) ? { ...e, subCategory: trimmed } : e;
          }
          return { ...e, items: e.items.map((it) => (it.mainCategory === main && it.subCategory === subModal.oldName) ? { ...it, subCategory: trimmed } : it) };
        }));
      }
    }
    setSubModal(null);
  };

  const removeSub = (main, subName) => {
    const used = countSub(main, subName);
    if (used > 0) {
      alert(`ลบไม่ได้ — มีค่าใช้จ่าย ${used} รายการที่ใช้หมวดหมู่ย่อยนี้อยู่ กรุณาเปลี่ยนหมวดหมู่ของรายการนั้นก่อน`);
      return;
    }
    setExpenseCategories({ ...expenseCategories, [main]: subsOf(main).filter((s) => s.name !== subName) });
  };

  return (
    <div>
      <Header title="หมวดหมู่ค่าใช้จ่าย" subtitle="ฐานข้อมูลหมวดหมู่ใหญ่/ย่อยของค่าใช้จ่าย ใช้เลือกตอนบันทึกค่าใช้จ่าย">
        <button style={btnPrimary} onClick={openAddMain}><Plus size={16} /> เพิ่มหมวดหมู่ใหญ่</button>
      </Header>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาหมวดหมู่ใหญ่หรือย่อย..." />

      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        {filtered.map((main) => (
          <div key={main} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "hidden" }}>
            <div style={{ background: "#f3f4f6", padding: "12px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>{main}</span>
                <span style={{ fontSize: 12, color: "#9ca3af" }}>{countMain(main)} รายการที่ใช้หมวดหมู่นี้</span>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button style={iconBtn} onClick={() => openAddSub(main)}><Plus size={14} /> เพิ่มหมวดหมู่ย่อย</button>
                <button style={iconBtn} onClick={() => openEditMain(main)}><Edit2 size={14} /> แก้ไข</button>
                <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบหมวดหมู่ใหญ่ "${main}" และหมวดหมู่ย่อยทั้งหมดในนี้ใช่หรือไม่?`, () => removeMain(main))}><Trash2 size={14} /> ลบ</button>
              </div>
            </div>
            <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
              <tbody>
                {subsOf(main).map((sub) => (
                  <tr key={sub.name}>
                    <td style={{ ...tdStyle, paddingLeft: 32, color: "#374151" }}>{sub.name}</td>
                    <td style={{ ...tdStyle, color: "#9ca3af", fontSize: 12 }}>{countSub(main, sub.name)} รายการ</td>
                    <td style={{ ...tdStyle, fontSize: 12 }}>
                      {Number(sub.openingBalance) > 0 ? (
                        <span style={{ background: "#faeeda", color: "#854f0b", padding: "2px 8px", borderRadius: 4, fontWeight: 600 }}>
                          ยอดยกมา ฿{fmt(sub.openingBalance)} ({monthLabelOf(sub.openingMonth)})
                        </span>
                      ) : (
                        <span style={{ color: "#d1d5db" }}>ไม่มียอดยกมา</span>
                      )}
                    </td>
                    <td style={{ ...tdStyle, textAlign: "right" }}>
                      <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                        <button style={iconBtn} onClick={() => openEditSub(main, sub)}><Edit2 size={14} /> แก้ไข</button>
                        <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบหมวดหมู่ย่อย "${sub.name}" ใช่หรือไม่?`, () => removeSub(main, sub.name))}><Trash2 size={14} /> ลบ</button>
                      </div>
                    </td>
                  </tr>
                ))}
                {subsOf(main).length === 0 && (
                  <tr><td colSpan={4} style={{ ...tdStyle, paddingLeft: 32, color: "#9ca3af" }}>ยังไม่มีหมวดหมู่ย่อย</td></tr>
                )}
              </tbody>
            </table>
            </div>
          </div>
        ))}
        {filtered.length === 0 && <p style={{ color: "#9ca3af", textAlign: "center", padding: 30 }}>ไม่พบหมวดหมู่ค่าใช้จ่าย</p>}
      </div>

      {mainModal && (
        <Modal title={mainModal.mode === "add" ? "เพิ่มหมวดหมู่ใหญ่" : "แก้ไขหมวดหมู่ใหญ่"} onClose={() => setMainModal(null)}>
          <Field label="ชื่อหมวดหมู่ใหญ่">
            <input style={inputStyle} value={mainName} onChange={(e) => setMainName(e.target.value)} placeholder="เช่น ค่าใช้จ่าย, ภาษี" />
          </Field>
          {mainModal.mode === "edit" && (
            <p style={{ fontSize: 12, color: "#9ca3af", margin: "-4px 0 8px" }}>
              * ถ้าเปลี่ยนชื่อ ค่าใช้จ่ายทุกรายการที่ใช้หมวดหมู่นี้จะถูกเปลี่ยนชื่อตามไปด้วยอัตโนมัติ
            </p>
          )}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setMainModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={saveMain}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {subModal && (
        <Modal title={subModal.mode === "add" ? `เพิ่มหมวดหมู่ย่อยใน "${subModal.main}"` : `แก้ไขหมวดหมู่ย่อยใน "${subModal.main}"`} onClose={() => setSubModal(null)}>
          <Field label="ชื่อหมวดหมู่ย่อย">
            <input style={inputStyle} value={subForm.name} onChange={(e) => setSubForm({ ...subForm, name: e.target.value })} placeholder="เช่น ค่าน้ำมัน/ขนส่ง" />
          </Field>
          {subModal.mode === "edit" && (
            <p style={{ fontSize: 12, color: "#9ca3af", margin: "-4px 0 8px" }}>
              * ถ้าเปลี่ยนชื่อ ค่าใช้จ่ายทุกรายการที่ใช้หมวดหมู่ย่อยนี้จะถูกเปลี่ยนชื่อตามไปด้วยอัตโนมัติ
            </p>
          )}
          <div style={{ background: "#faeeda", borderRadius: 8, padding: "12px 16px", marginTop: 8 }}>
            <Field label="ยอดยกมา (บาท)">
              <input type="number" min={0} style={inputStyle} value={subForm.openingBalance} onChange={(e) => setSubForm({ ...subForm, openingBalance: e.target.value })} placeholder="0" />
            </Field>
            <Field label="ของเดือน">
              <input type="month" style={inputStyle} value={subForm.openingMonth} onChange={(e) => setSubForm({ ...subForm, openingMonth: e.target.value })} />
            </Field>
            <p style={{ fontSize: 11, color: "#854f0b", margin: 0 }}>
              * ยอดสะสมก่อนเริ่มใช้ระบบ กรอกครั้งเดียว — จะรวมเข้าไปเฉพาะตอนดูแดชบอร์ด/รายงานของเดือนที่ระบุไว้เท่านั้น
            </p>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setSubModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={saveSub}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ===================================================================
// STORE BANK ACCOUNTS TAB (บัญชีธนาคารของร้าน)
// ===================================================================
function StoreBankAccountsTab({ accounts, setAccounts, purchases, sales, expenses, deposits, bankTransfers, customers }) {
  const [modal, setModal] = useState(null);
  const [statementModal, setStatementModal] = useState(null); // {account}
  const [stmtYear, setStmtYear] = useState(new Date().getFullYear());
  const [stmtMonth, setStmtMonth] = useState(new Date().getMonth() + 1);
  const [stmtMode, setStmtMode] = useState("month"); // "month" | "range"
  const [stmtDateFrom, setStmtDateFrom] = useState(new Date().toISOString().slice(0, 10));
  const [stmtDateTo, setStmtDateTo] = useState(new Date().toISOString().slice(0, 10));
  const blank = { id: "", bankName: BANK_NAMES[0], accountNo: "", accountName: "", branch: "", openingBalance: 0, accountType: "" };
  const [form, setForm] = useState(blank);

  const openAdd = () => { setForm({ ...blank, id: "SB" + Date.now().toString().slice(-6) }); setModal({ mode: "add" }); };
 const openEdit = (item) => { setForm({ openingBalance: 0, accountType: "", ...item }); setModal({ mode: "edit", item }); };
  const save = () => {
    if (!form.bankName || !form.accountNo.trim()) return;
    const cleaned = { ...form, openingBalance: Number(form.openingBalance) || 0 };
    if (modal.mode === "add") setAccounts([...accounts, cleaned]);
    else setAccounts(accounts.map((a) => (a.id === modal.item.id ? cleaned : a)));
    setModal(null);
  };

  const remove = (id) => setAccounts(accounts.filter((a) => a.id !== id));

  // สร้าง Statement รายการเดินบัญชี
  const buildStatement = (acc) => {
    const startDate = stmtMode === "range" ? stmtDateFrom
      : `${stmtYear}-${String(stmtMonth).padStart(2,"0")}-01`;
    const endDate = stmtMode === "range" ? stmtDateTo
      : `${stmtYear}-${String(stmtMonth).padStart(2,"0")}-${String(new Date(stmtYear, stmtMonth, 0).getDate()).padStart(2,"0")}`;
    const inRange   = (d) => d >= startDate && d <= endDate;
    const rows = [];

    // รายรับ: ใบขายที่ชำระเข้าบัญชีนี้
    (sales || []).forEach((inv) => {
      (inv.payments || []).forEach((p) => {
        if (p.toStoreBankId === acc.id && inRange(p.date)) {
          rows.push({ date: p.date, type: "รับชำระ", ref: inv.id, description: `รับชำระ Invoice ${inv.id}`, credit: Number(p.amount) || 0, debit: 0 });
        }
      });
    });

    // รายจ่าย: ค่าใช้จ่ายที่จ่ายจากบัญชีนี้
    (expenses || []).forEach((e) => {
      (e.payments || []).forEach((p) => {
        if (p.fromStoreBankId === acc.id && inRange(p.date || e.billDate || e.date)) {
          rows.push({ date: p.date || e.billDate || e.date, type: "จ่ายค่าใช้จ่าย", ref: e.refNo || e.id, description: `ค่าใช้จ่าย ${e.refNo || e.id}`, debit: Number(p.amount) || 0, credit: 0 });
        }
      });
      // กรณีบันทึกบัญชีตรง (ไม่ผ่าน payments)
      if (!(e.payments && e.payments.length > 0) && e.fromStoreBankId === acc.id && inRange(e.billDate || e.date)) {
        const items = (e.items && e.items.length > 0) ? e.items : [{ amount: e.amount }];
        const amt = items.reduce((s, it) => s + (Number(it.amount) || 0), 0);
        rows.push({ date: e.billDate || e.date, type: "จ่ายค่าใช้จ่าย", ref: e.refNo || e.id, description: `ค่าใช้จ่าย ${e.refNo || e.id}`, debit: amt, credit: 0 });
      }
    });

    // รายจ่าย: ใบรับสินค้า (จ่ายเงินให้ลูกค้า)
    (purchases || []).forEach((po) => {
      (po.payments || []).forEach((p) => {
        if (p.fromStoreBankId === acc.id && inRange(p.date)) {
          rows.push({ date: p.date, type: "จ่ายรับสินค้า", ref: po.id, description: `จ่ายรับสินค้า ${po.id}`, debit: Number(p.amount) || 0, credit: 0 });
        }
      });
    });

    // รายจ่าย: เงินมัดจำที่จ่ายล่วงหน้าให้ลูกค้าจากบัญชีนี้
    (deposits || []).forEach((d) => {
      if (d.fromStoreBankId === acc.id && inRange(d.date)) {
        const cust = (customers || []).find((c) => c.id === d.customerId);
        const custLabel = cust ? cust.name : (d.customerId || "ลูกค้า");
        rows.push({ date: d.date, type: "จ่ายมัดจำ", ref: d.id, description: `จ่ายมัดจำให้ ${custLabel}${d.note ? " — " + d.note : ""}`, debit: Number(d.amount) || 0, credit: 0 });
      }
    });

    // โยกเงินระหว่างบัญชี: เข้า (รับโอน) และ ออก (โอนออก)
    (bankTransfers || []).forEach((t) => {
      if (t.toBankId === acc.id && inRange(t.date)) {
        const fromAcc = (accounts || []).find((b) => b.id === t.fromBankId);
        const fromLabel = fromAcc ? `${fromAcc.bankName} ${fromAcc.accountNo}` : "บัญชีอื่น";
        rows.push({ date: t.date, type: "รับโอนเงิน", ref: t.id, description: `รับโอนจาก ${fromLabel}${t.note ? " — " + t.note : ""}`, credit: Number(t.amount) || 0, debit: 0 });
      }
      if (t.fromBankId === acc.id && inRange(t.date)) {
        const toAcc = (accounts || []).find((b) => b.id === t.toBankId);
        const toLabel = toAcc ? `${toAcc.bankName} ${toAcc.accountNo}` : "บัญชีอื่น";
        rows.push({ date: t.date, type: "โอนเงินออก", ref: t.id, description: `โอนไป ${toLabel}${t.note ? " — " + t.note : ""}`, debit: Number(t.amount) || 0, credit: 0 });
      }
    });

    rows.sort((a, b) => a.date.localeCompare(b.date));

    // คำนวณยอดคงเหลือ
    let balance = Number(acc.openingBalance) || 0;
    const withBalance = rows.map((r) => {
      balance += r.credit - r.debit;
      return { ...r, balance };
    });
    return { rows: withBalance, startBalance: Number(acc.openingBalance) || 0, endBalance: balance };
  };

  const MONTH_NAMES = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  const yearOptions = [];
  for (let y = 2024; y <= new Date().getFullYear() + 2; y++) yearOptions.push(y);

  return (
    <div>
      <Header title="บัญชีธนาคารของร้าน" subtitle="จัดการบัญชีธนาคาร — ดูรายการเดินบัญชี (Statement) ได้">
        <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> เพิ่มบัญชีธนาคาร</button>
      </Header>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))", gap: 14 }}>
        {[...accounts].reverse().map((a) => (
          <div key={a.id} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "16px 18px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div style={{ width: 36, height: 36, borderRadius: 8, background: a.accountType === "เงินสด" ? "#e1f5ee" : "#e6f1fb", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  {a.accountType === "เงินสด" ? <Wallet size={18} color="#0f6e56" /> : <Landmark size={18} color="#185fa5" />}
                </div>
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <div style={{ fontWeight: 600, fontSize: 14 }}>{a.bankName}</div>
                    {a.accountType ? (
                      <span style={{
                        fontSize: 10, fontWeight: 600, padding: "1px 7px", borderRadius: 5,
                        background: a.accountType === "เงินสด" ? "#e1f5ee" : "#e6f1fb",
                        color: a.accountType === "เงินสด" ? "#0f6e56" : "#185fa5",
                      }}>{a.accountType}</span>
                    ) : (
                      <span style={{ fontSize: 10, fontWeight: 600, padding: "1px 7px", borderRadius: 5, background: "#faeeda", color: "#854f0b" }}>ยังไม่ระบุประเภท</span>
                    )}
                  </div>
                  <div style={{ fontFamily: "'JetBrains Mono', monospace", fontSize: 13, color: "#6b7280" }}>{a.accountNo}</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button style={iconBtn} onClick={() => openEdit(a)}><Edit2 size={14} /></button>
                <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบบัญชี "${a.bankName} ${a.accountNo}" ใช่หรือไม่?`, () => remove(a.id))}><Trash2 size={14} /></button>
              </div>
            </div>
            <div style={{ fontSize: 13, color: "#6b7280" }}>
              <div>ชื่อบัญชี: {a.accountName || "-"}</div>
              <div>สาขา: {a.branch || "-"}</div>
            </div>
            {(Number(a.openingBalance) > 0) && (
              <div style={{ marginTop: 6, padding: "5px 10px", background: "#e6f1fb", borderRadius: 6, fontSize: 12, color: "#185fa5", fontWeight: 600 }}>
                ยอดยกมา: ฿{fmt(a.openingBalance)}
              </div>
            )}
            <button
              style={{ ...btnSecondary, marginTop: 10, width: "100%", fontSize: 12 }}
              onClick={() => setStatementModal(a)}
            >
              <History size={13} /> ดูรายการเดินบัญชี (Statement)
            </button>
          </div>
        ))}
        {accounts.length === 0 && <p style={{ color: "#9ca3af" }}>ยังไม่มีบัญชีธนาคารของร้าน</p>}
      </div>

      {/* Statement Modal */}
      {statementModal && (() => {
        const stmt = buildStatement(statementModal);
        const totalCredit = stmt.rows.reduce((s, r) => s + r.credit, 0);
        const totalDebit  = stmt.rows.reduce((s, r) => s + r.debit, 0);
        return (
          <Modal title={`Statement — ${statementModal.bankName} ${statementModal.accountNo}`} onClose={() => setStatementModal(null)} wide fullscreen>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
              {/* mode toggle */}
              <div style={{ display: "flex", borderRadius: 8, overflow: "hidden", border: "1px solid #d1d5db" }}>
                {[{key:"month",label:"รายเดือน"},{key:"range",label:"เลือกช่วงวันที่"}].map((opt) => (
                  <button key={opt.key} onClick={() => setStmtMode(opt.key)}
                    style={{ padding:"7px 14px", border:"none", cursor:"pointer", fontSize:13, fontWeight:600,
                      background: stmtMode===opt.key ? "#0c443c" : "#fff",
                      color: stmtMode===opt.key ? "#fff" : "#6b7280" }}>
                    {opt.label}
                  </button>
                ))}
              </div>
              {stmtMode === "month" && <>
                <select style={{ ...inputStyle, width: 140 }} value={stmtMonth} onChange={(e) => setStmtMonth(Number(e.target.value))}>
                  {MONTH_NAMES.slice(1).map((n, i) => <option key={i+1} value={i+1}>{n}</option>)}
                </select>
                <select style={{ ...inputStyle, width: 100 }} value={stmtYear} onChange={(e) => setStmtYear(Number(e.target.value))}>
                  {yearOptions.map((y) => <option key={y} value={y}>ปี {y}</option>)}
                </select>
              </>}
              {stmtMode === "range" && <>
                <input type="date" style={{ ...inputStyle, width: 160 }} value={stmtDateFrom} onChange={(e) => setStmtDateFrom(e.target.value)} />
                <span style={{ fontSize: 13, color: "#6b7280" }}>ถึง</span>
                <input type="date" style={{ ...inputStyle, width: 160 }} value={stmtDateTo} onChange={(e) => setStmtDateTo(e.target.value)} />
              </>}
              <button style={btnSecondary} onClick={() => printAsPDF("stmt-print", `Statement ${statementModal.accountNo}`)}>
                <Download size={14} /> พิมพ์
              </button>
              <button style={btnSecondary} onClick={() => {
                const rows = [
                  [`Statement — ${statementModal.bankName} ${statementModal.accountNo}`, "", "", "", "", ""],
                  [stmtMode === "range" ? `${stmtDateFrom} ถึง ${stmtDateTo}` : `${MONTH_NAMES[stmtMonth]} ${stmtYear}`, "", "", "", "", ""],
                  ["วันที่", "ประเภท", "อ้างอิง", "รายการ", "ฝาก (เข้า)", "ถอน (ออก)", "คงเหลือ"],
                  ["ยอดยกมา", "", "", "", "", "", stmt.startBalance],
                  ...stmt.rows.map(r => [r.date, r.type, r.ref, r.description, r.credit || "", r.debit || "", r.balance]),
                  ["", "", "", "รวม", totalCredit, totalDebit, stmt.endBalance],
                ];
                exportExcel(rows, `Statement_${statementModal.accountNo}_${stmtYear}${String(stmtMonth).padStart(2,"0")}.xlsx`, "Statement");
              }}>
                <FileSpreadsheet size={14} /> Excel
              </button>
            </div>

            <div id="stmt-print">
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 16 }}>
                <div style={{ background: "#f9fafb", borderRadius: 8, padding: "10px 14px", fontSize: 13 }}>
                  <div style={{ color: "#6b7280", marginBottom: 2 }}>ยอดยกมา</div>
                  <div style={{ fontWeight: 700, color: "#185fa5" }}>฿{fmt(stmt.startBalance)}</div>
                </div>
                <div style={{ background: "#e1f5ee", borderRadius: 8, padding: "10px 14px", fontSize: 13 }}>
                  <div style={{ color: "#6b7280", marginBottom: 2 }}>รับเข้ารวม</div>
                  <div style={{ fontWeight: 700, color: "#0f6e56" }}>฿{fmt(totalCredit)}</div>
                </div>
                <div style={{ background: "#faece7", borderRadius: 8, padding: "10px 14px", fontSize: 13 }}>
                  <div style={{ color: "#6b7280", marginBottom: 2 }}>จ่ายออกรวม</div>
                  <div style={{ fontWeight: 700, color: "#993c1d" }}>฿{fmt(totalDebit)}</div>
                </div>
              </div>

              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr>
                    <th style={thStyle}>วันที่</th>
                    <th style={thStyle}>ประเภท</th>
                    <th style={thStyle}>เลขอ้างอิง</th>
                    <th style={thStyle}>รายการ</th>
                    <th style={{ ...thStyle, textAlign: "right", color: "#0f6e56" }}>ฝาก (เข้า)</th>
                    <th style={{ ...thStyle, textAlign: "right", color: "#993c1d" }}>ถอน (ออก)</th>
                    <th style={{ ...thStyle, textAlign: "right" }}>คงเหลือ</th>
                  </tr>
                </thead>
                <tbody>
                  <tr style={{ background: "#f9fafb" }}>
                    <td style={tdStyle} colSpan={6}><strong>ยอดยกมา</strong></td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#185fa5" }}>฿{fmt(stmt.startBalance)}</td>
                  </tr>
                  {stmt.rows.map((r, i) => (
                    <tr key={i}>
                      <td style={tdStyle}>{r.date}</td>
                      <td style={tdStyle}><span style={{ background: r.credit > 0 ? "#e1f5ee" : "#faece7", color: r.credit > 0 ? "#0f6e56" : "#993c1d", padding: "2px 8px", borderRadius: 4, fontSize: 11, fontWeight: 500 }}>{r.type}</span></td>
                      <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 11 }}>{r.ref}</td>
                      <td style={tdStyle}>{r.description}</td>
                      <td style={{ ...tdStyle, textAlign: "right", color: "#0f6e56", fontWeight: r.credit > 0 ? 600 : 400 }}>{r.credit > 0 ? `฿${fmt(r.credit)}` : "-"}</td>
                      <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d", fontWeight: r.debit > 0 ? 600 : 400 }}>{r.debit > 0 ? `฿${fmt(r.debit)}` : "-"}</td>
                      <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: r.balance >= 0 ? "#1f2937" : "#a32d2d" }}>฿{fmt(r.balance)}</td>
                    </tr>
                  ))}
                  {stmt.rows.length === 0 && (
                    <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีรายการในเดือนนี้</td></tr>
                  )}
                </tbody>
                <tfoot>
                  <tr style={{ background: "#f3f4f6" }}>
                    <td colSpan={4} style={{ ...tdStyle, fontWeight: 700 }}>รวม / ยอดคงเหลือสิ้นเดือน</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>฿{fmt(totalCredit)}</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>฿{fmt(totalDebit)}</td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 15, color: stmt.endBalance >= 0 ? "#185fa5" : "#a32d2d" }}>฿{fmt(stmt.endBalance)}</td>
                  </tr>
                </tfoot>
              </table>
              {/* ตารางสรุปเงินมัดจำ */}
              {(() => {
                const depositRows = customers.map((c) => {
                  const opening = Number(c.depositOpening) || 0;
                  const given = (deposits || []).filter((d) => d.customerId === c.id).reduce((s, d) => s + (Number(d.amount) || 0), 0);
                  const used = purchases.reduce((s, po) => s + (po.payments || []).filter((p) => p.fromStoreBankId === "DEPOSIT" && po.customerId === c.id).reduce((s2, p) => s2 + (Number(p.amount) || 0), 0), 0);
                  const remaining = opening + given - used;
                  return { name: c.name, opening, given, used, remaining };
                }).filter((r) => r.opening > 0 || r.given > 0);
                if (depositRows.length === 0) return null;
                const totalRemaining = depositRows.reduce((s, r) => s + r.remaining, 0);
                return (
                  <div style={{ marginTop: 20 }}>
                    <div style={{ background: "#854f0b", color: "#fff", padding: "10px 16px", fontWeight: 700, fontSize: 14, borderRadius: "8px 8px 0 0" }}>
                      สรุปเงินมัดจำคงเหลือต่อลูกค้า
                    </div>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                      <thead>
                        <tr>
                          <th style={thStyle}>ลูกค้า</th>
                          <th style={{ ...thStyle, textAlign: "right" }}>ยอดยกมา</th>
                          <th style={{ ...thStyle, textAlign: "right", color: "#0f6e56" }}>จ่ายมัดจำเพิ่ม</th>
                          <th style={{ ...thStyle, textAlign: "right", color: "#993c1d" }}>หักไปแล้ว</th>
                          <th style={{ ...thStyle, textAlign: "right" }}>คงเหลือ</th>
                        </tr>
                      </thead>
                      <tbody>
                        {depositRows.map((r, i) => (
                          <tr key={i}>
                            <td style={tdStyle}>{r.name}</td>
                            <td style={{ ...tdStyle, textAlign: "right", color: "#6b7280" }}>฿{fmt(r.opening)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", color: "#0f6e56" }}>+฿{fmt(r.given)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", color: "#993c1d" }}>-฿{fmt(r.used)}</td>
                            <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: r.remaining > 0 ? "#854f0b" : "#6b7280" }}>฿{fmt(r.remaining)}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr style={{ background: "#faeeda" }}>
                          <td colSpan={4} style={{ ...tdStyle, fontWeight: 700, color: "#854f0b" }}>รวมเงินมัดจำคงเหลือทั้งหมด</td>
                          <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, fontSize: 15, color: "#854f0b" }}>฿{fmt(totalRemaining)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                );
              })()}
            </div>
          </Modal>
        );
      })()}

      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มบัญชีธนาคารของร้าน" : "แก้ไขบัญชีธนาคารของร้าน"} onClose={() => setModal(null)}>
          <Field label="ประเภทบัญชี">
            <div style={{ display: "flex", gap: 10 }}>
              {["ธนาคาร", "เงินสด"].map((t) => (
                <button
                  key={t}
                  type="button"
                  onClick={() => setForm({ ...form, accountType: t })}
                  style={{
                    flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
                    padding: "9px 14px", borderRadius: 8, fontSize: 14, cursor: "pointer",
                    border: form.accountType === t ? "1.5px solid #1d9e75" : "1px solid #d1d5db",
                    background: form.accountType === t ? "#e1f5ee" : "#fff",
                    color: form.accountType === t ? "#0c443c" : "#6b7280",
                    fontWeight: form.accountType === t ? 600 : 400,
                  }}
                >
                  {t === "ธนาคาร" ? <Landmark size={15} /> : <Wallet size={15} />} {t}
                </button>
              ))}
            </div>
            {!form.accountType && <p style={{ fontSize: 11, color: "#854f0b", marginTop: 4, marginBottom: 0 }}>* กรุณาเลือกประเภทบัญชี เพื่อให้แสดงผลถูกกลุ่มในแดชบอร์ด</p>}
          </Field>
          <Field label="ธนาคาร">
  <input style={inputStyle} list="bank-name-options" value={form.bankName} onChange={(e) => setForm({ ...form, bankName: e.target.value })} placeholder="เลือกหรือพิมพ์ชื่อธนาคาร" />
  <datalist id="bank-name-options">
    {BANK_NAMES.map((n) => <option key={n} value={n} />)}
  </datalist>
</Field>
          <Field label="เลขที่บัญชี"><input style={inputStyle} value={form.accountNo} onChange={(e) => setForm({ ...form, accountNo: e.target.value })} /></Field>
          <Field label="ชื่อบัญชี"><input style={inputStyle} value={form.accountName} onChange={(e) => setForm({ ...form, accountName: e.target.value })} /></Field>
          <Field label="สาขา"><input style={inputStyle} value={form.branch} onChange={(e) => setForm({ ...form, branch: e.target.value })} /></Field>
          <div style={{ background: "#e6f1fb", borderRadius: 8, padding: "12px 16px", marginTop: 8 }}>
            <Field label="ยอดคงเหลือยกมา (บาท)">
              <input type="number" min={0} style={inputStyle} value={form.openingBalance} onChange={(e) => setForm({ ...form, openingBalance: e.target.value })} placeholder="0" />
            </Field>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}
function BankTransferTab({ storeBankAccounts, bankTransfers, setBankTransfers }) {
  const transfers = bankTransfers || [];
  const [modal, setModal] = useState(null);
  const blankForm = () => ({
    id: "TF" + Date.now().toString().slice(-6),
    date: new Date().toISOString().slice(0, 10),
    fromBankId: storeBankAccounts[0]?.id || "",
    toBankId: storeBankAccounts[1]?.id || storeBankAccounts[0]?.id || "",
    amount: 0,
    note: "",
  });
  const [form, setForm] = useState(blankForm());
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");

  const bankName = (id) => {
    const b = storeBankAccounts.find((b) => b.id === id);
    return b ? `${b.bankName} ${b.accountNo}` : "-";
  };

  const save = () => {
    if (!form.fromBankId || !form.toBankId || !(Number(form.amount) > 0)) return;
    if (form.fromBankId === form.toBankId) { alert("บัญชีต้นทางและปลายทางต้องต่างกัน"); return; }
    const t = { ...form, amount: Number(form.amount) };
    if (modal.mode === "add") setBankTransfers([t, ...transfers]);
    else setBankTransfers(transfers.map((x) => x.id === modal.item.id ? t : x));
    setModal(null);
  };

  const totalOut = (bankId) => transfers.filter((t) => t.fromBankId === bankId).reduce((s, t) => s + t.amount, 0);
  const totalIn = (bankId) => transfers.filter((t) => t.toBankId === bankId).reduce((s, t) => s + t.amount, 0);

  return (
    <div>
      <Header title="โยกเงินระหว่างธนาคาร" subtitle="บันทึกการโอนเงินระหว่างบัญชีธนาคารของร้าน">
        <button style={btnPrimary} onClick={() => { setForm(blankForm()); setModal({ mode: "add" }); }}><Plus size={16} /> บันทึกโยกเงิน</button>
      </Header>

      {storeBankAccounts.length < 2 && (
        <div style={{ background: "#faeeda", border: "1px solid #f0c070", borderRadius: 10, padding: "12px 16px", marginBottom: 16, fontSize: 13, color: "#854f0b" }}>
          ⚠️ ต้องมีบัญชีธนาคารร้านอย่างน้อย 2 บัญชีเพื่อโยกเงิน — ไปที่ "บัญชีธนาคารร้าน" เพื่อเพิ่ม
        </div>
      )}

      {/* สรุปยอดโยกเงินต่อบัญชี */}
      {storeBankAccounts.length > 0 && transfers.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 12, marginBottom: 20 }}>
          {storeBankAccounts.map((b) => (
            <div key={b.id} style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "14px 16px" }}>
              <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 6 }}>{b.bankName} {b.accountNo}</div>
              <div style={{ fontSize: 12, color: "#0f6e56" }}>รับโอนเข้า: ฿{fmt(totalIn(b.id))}</div>
              <div style={{ fontSize: 12, color: "#993c1d" }}>โอนออก: ฿{fmt(totalOut(b.id))}</div>
            </div>
          ))}
        </div>
      )}

      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาเลขที่, บัญชี..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />
        {transfers.filter((t) => (t.id || "").includes(search) && (!dateFrom || (t.date || "") >= dateFrom) && (!dateTo || (t.date || "") <= dateTo)).map((t) => (
          <div key={t.id} style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "14px 18px", display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
            <div style={{ flex: 1 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                <span style={{ fontFamily: "monospace", fontSize: 12, color: "#9ca3af" }}>{t.id}</span>
                <span style={{ fontSize: 13, color: "#6b7280" }}>{t.date}</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 14 }}>
                <span style={{ fontWeight: 600, color: "#993c1d" }}>{bankName(t.fromBankId)}</span>
                <ArrowRight size={14} color="#9ca3af" />
                <span style={{ fontWeight: 600, color: "#0f6e56" }}>{bankName(t.toBankId)}</span>
              </div>
              {t.note && <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 4 }}>{t.note}</div>}
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: 11, color: "#9ca3af" }}>ยอดโอน</div>
              <div style={{ fontWeight: 700, fontSize: 18, color: "#185fa5" }}>฿{fmt(t.amount)}</div>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button style={iconBtn} onClick={() => { setForm({ ...t }); setModal({ mode: "edit", item: t }); }}><Edit2 size={14} /></button>
              <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบรายการโยกเงิน "${t.id}" จำนวน ฿${fmt(t.amount)} ใช่หรือไม่?`, () => setBankTransfers(transfers.filter((x) => x.id !== t.id)))}><Trash2 size={14} /></button>
            </div>
          </div>
        ))}
        {transfers.length === 0 && <div style={{ textAlign: "center", color: "#9ca3af", padding: 40 }}>ยังไม่มีรายการโยกเงิน</div>}
      </div>

      {modal && (
        <Modal title={modal.mode === "add" ? "บันทึกโยกเงิน" : "แก้ไขโยกเงิน"} onClose={() => setModal(null)}>
          <Field label="วันที่โอน"><input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
          <Field label="โอนจากบัญชี">
            <select style={inputStyle} value={form.fromBankId} onChange={(e) => setForm({ ...form, fromBankId: e.target.value })}>
              {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
            </select>
          </Field>
          <Field label="โอนไปบัญชี">
            <select style={inputStyle} value={form.toBankId} onChange={(e) => setForm({ ...form, toBankId: e.target.value })}>
              {storeBankAccounts.map((b) => <option key={b.id} value={b.id}>{b.bankName} {b.accountNo}</option>)}
            </select>
          </Field>
          <Field label="จำนวนเงิน (บาท)"><input type="number" min={0} style={inputStyle} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></Field>
          <Field label="หมายเหตุ"><input style={inputStyle} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ===================================================================
// ASSETS TAB (ทะเบียนทรัพย์สิน)
// ===================================================================
// ===================================================================
function AssetsTab({ assets, setAssets }) {
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");

  const ASSET_CATEGORIES = ["ยานพาหนะ", "เครื่องจักร/อุปกรณ์", "อาคาร/สิ่งปลูกสร้าง", "คอมพิวเตอร์/IT", "เฟอร์นิเจอร์/ของตกแต่ง", "อื่นๆ"];

  const blankForm = () => ({
    id: genId("AS", assets),
    name: "", category: ASSET_CATEGORIES[0], purchaseDate: new Date().toISOString().slice(0, 10),
    cost: 0, lifeYears: 5, depreciationMethod: "เส้นตรง", note: "",
  });
  const [form, setForm] = useState(blankForm());

  const annualDepreciation = (a) => a.depreciationMethod === "เส้นตรง" ? Number(a.cost) / Number(a.lifeYears) : 0;
  const monthlyDepreciation = (a) => annualDepreciation(a) / 12;
  const yearsUsed = (a) => {
    const ms = new Date() - new Date(a.purchaseDate);
    return ms / (1000 * 60 * 60 * 24 * 365.25);
  };
  const accumulatedDepreciation = (a) => Math.min(Number(a.cost), annualDepreciation(a) * yearsUsed(a));
  const bookValue = (a) => Math.max(0, Number(a.cost) - accumulatedDepreciation(a));

  const filtered = assets.filter((a) => a.name.includes(search) || a.category.includes(search) || a.id.includes(search)).filter((a) => (!dateFrom || (a.purchaseDate || "") >= dateFrom) && (!dateTo || (a.purchaseDate || "") <= dateTo))
    .sort((a, b) => (b.purchaseDate || "").localeCompare(a.purchaseDate || "") || (b.id || "").localeCompare(a.id || ""));

  const save = () => {
    if (!form.name.trim()) return;
    const cleaned = { ...form, cost: Number(form.cost) || 0, lifeYears: Number(form.lifeYears) || 1 };
    if (modal.mode === "add") setAssets([...assets, cleaned]);
    else setAssets(assets.map((a) => a.id === modal.item.id ? cleaned : a));
    setModal(null);
  };

  const totalCost = assets.reduce((s, a) => s + Number(a.cost), 0);
  const totalBookValue = assets.reduce((s, a) => s + bookValue(a), 0);
  const totalAccDep = assets.reduce((s, a) => s + accumulatedDepreciation(a), 0);

  return (
    <div>
      <Header title="ทะเบียนทรัพย์สิน" subtitle="บันทึกและคำนวณค่าเสื่อมราคาทรัพย์สินของร้าน">
        <button style={btnPrimary} onClick={() => { setForm(blankForm()); setModal({ mode: "add" }); }}><Plus size={16} /> เพิ่มทรัพย์สิน</button>
      </Header>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 20 }}>
        {[
          { label: "ราคาทุนรวมทั้งหมด", value: fmt(totalCost), color: "#185fa5", bg: "#e6f1fb" },
          { label: "ค่าเสื่อมราคาสะสม", value: fmt(totalAccDep), color: "#854f0b", bg: "#faeeda" },
          { label: "มูลค่าตามบัญชีรวม", value: fmt(totalBookValue), color: "#0f6e56", bg: "#e1f5ee" },
        ].map((c) => (
          <div key={c.label} style={{ background: c.bg, borderRadius: 12, padding: "14px 18px" }}>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 4 }}>{c.label}</div>
            <div style={{ fontWeight: 700, fontSize: 20, color: c.color }}>฿{c.value}</div>
          </div>
        ))}
      </div>

      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาชื่อทรัพย์สิน, หมวดหมู่..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />

      <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 760 }}>
          <thead>
            <tr>
              <th style={thStyle}>รหัส</th>
              <th style={thStyle}>ชื่อทรัพย์สิน</th>
              <th style={thStyle}>หมวดหมู่</th>
              <th style={thStyle}>วันที่ซื้อ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ราคาทุน</th>
              <th style={{ ...thStyle, textAlign: "right" }}>อายุ (ปี)</th>
              <th style={{ ...thStyle, textAlign: "right" }}>เสื่อม/ปี</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ค่าเสื่อมสะสม</th>
              <th style={{ ...thStyle, textAlign: "right" }}>มูลค่าตามบัญชี</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((a) => (
              <tr key={a.id}>
                <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{a.id}</td>
                <td style={{ ...tdStyle, fontWeight: 600 }}>{a.name}</td>
                <td style={tdStyle}><Badge text={a.category} /></td>
                <td style={tdStyle}>{a.purchaseDate}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(a.cost)}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>{a.lifeYears}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(annualDepreciation(a))}</td>
                <td style={{ ...tdStyle, textAlign: "right", color: "#854f0b" }}>{fmt(accumulatedDepreciation(a))}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#0f6e56" }}>{fmt(bookValue(a))}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <button style={iconBtn} onClick={() => { setForm({ ...a }); setModal({ mode: "edit", item: a }); }}><Edit2 size={14} /></button>
                    <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบทรัพย์สิน "${a.name}" ใช่หรือไม่?`, () => setAssets(assets.filter((x) => x.id !== a.id)))}><Trash2 size={14} /></button>
                  </div>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={10} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่พบทรัพย์สิน</td></tr>}
          </tbody>
          {assets.length > 0 && (
            <tfoot>
              <tr>
                <td colSpan={4} style={{ ...tdStyle, fontWeight: 700 }}>รวม</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(totalCost)}</td>
                <td style={tdStyle}></td>
                <td style={tdStyle}></td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#854f0b" }}>{fmt(totalAccDep)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>{fmt(totalBookValue)}</td>
                <td style={tdStyle}></td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>

      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มทรัพย์สิน" : "แก้ไขทรัพย์สิน"} onClose={() => setModal(null)} wide>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            <Field label="รหัสทรัพย์สิน"><input style={inputStyle} value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} disabled={modal.mode === "edit"} /></Field>
            <Field label="ชื่อทรัพย์สิน"><input style={inputStyle} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></Field>
            <Field label="หมวดหมู่">
              <select style={inputStyle} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>
                {ASSET_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="วันที่ซื้อ"><input type="date" style={inputStyle} value={form.purchaseDate} onChange={(e) => setForm({ ...form, purchaseDate: e.target.value })} /></Field>
            <Field label="ราคาทุน (บาท)"><input type="number" min={0} style={inputStyle} value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} /></Field>
            <Field label="อายุการใช้งาน (ปี)"><input type="number" min={1} max={50} style={inputStyle} value={form.lifeYears} onChange={(e) => setForm({ ...form, lifeYears: e.target.value })} /></Field>
            <Field label="วิธีเสื่อมราคา">
              <select style={inputStyle} value={form.depreciationMethod} onChange={(e) => setForm({ ...form, depreciationMethod: e.target.value })}>
                <option value="เส้นตรง">เส้นตรง (Straight-Line)</option>
                <option value="ยอดคงเหลือลดลง">ยอดคงเหลือลดลง</option>
              </select>
            </Field>
          </div>
          <Field label="หมายเหตุ"><input style={inputStyle} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
          {Number(form.cost) > 0 && Number(form.lifeYears) > 0 && (
            <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", marginTop: 8, fontSize: 13 }}>
              <Row label="ค่าเสื่อมราคาต่อปี" value={`฿${fmt(Number(form.cost) / Number(form.lifeYears))}`} />
              <Row label="ค่าเสื่อมราคาต่อเดือน" value={`฿${fmt(Number(form.cost) / Number(form.lifeYears) / 12)}`} />
            </div>
          )}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}
    </div>
  );
}


// ===================================================================
// COMPANY SETTINGS TAB (ตั้งค่าร้าน / โลโก้)
// ===================================================================
function CompanySettingsTab({ settings, setSettings, shopProfile, setShopProfile }) {
  const cs = settings || {};
  const sp = shopProfile || {};
  const set = (field, value) => setSettings((prev) => ({ ...prev, [field]: value }));
  const setSP = (field, value) => setShopProfile((prev) => ({ ...prev, [field]: value }));
  const [saved, setSaved] = useState(false);

  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) { alert("รูปภาพต้องไม่เกิน 2MB"); return; }
    const reader = new FileReader();
    reader.onload = (ev) => set("logo", ev.target.result);
    reader.readAsDataURL(file);
  };

  const handleSidebarLogoUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (file.size > 2 * 1024 * 1024) { alert("รูปภาพต้องไม่เกิน 2MB"); return; }
    const reader = new FileReader();
    reader.onload = (ev) => setSP("logo", ev.target.result);
    reader.readAsDataURL(file);
  };

  const handleSave = () => {
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };

  const sCard = { background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "20px 24px", marginBottom: 16 };

  return (
    <div>
      <Header title="ตั้งค่ากิจการ" subtitle="แยกเป็น 2 ส่วน — โปรไฟล์หน้าแอป (sidebar) และข้อมูลเอกสาร/บิล">
        <button style={btnPrimary} onClick={handleSave}>
          {saved ? <><CheckCircle2 size={16} /> บันทึกแล้ว!</> : <><Save size={16} /> บันทึก</>}
        </button>
      </Header>

      {/* ===== ส่วนที่ 1: โปรไฟล์ Sidebar ===== */}
      <div style={{ ...sCard, border: "2px solid #1d9e75" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <div style={{ width: 28, height: 28, borderRadius: 6, background: "#0c443c", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Boxes size={14} color="#9fe1cb" />
          </div>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#0c443c" }}>โปรไฟล์แอป (แสดงในแถบเมนูซ้าย)</h3>
          <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 4 }}>แยกอิสระจากข้อมูลบิล</span>
        </div>

        <div style={{ display: "flex", gap: 20, alignItems: "flex-start", flexWrap: "wrap" }}>
          {/* sidebar logo preview */}
          <div style={{ flexShrink: 0 }}>
            <div style={{ background: "#0c443c", borderRadius: 10, padding: "12px 14px", display: "flex", alignItems: "center", gap: 10, width: 200 }}>
              {sp.logo ? (
                <img src={sp.logo} alt="logo" style={{ width: 38, height: 38, borderRadius: 8, objectFit: "contain", background: "#fff", padding: 3 }} />
              ) : (
                <div style={{ width: 38, height: 38, borderRadius: 8, background: "#1d9e75", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <Boxes size={18} color="#04342c" />
                </div>
              )}
              <div>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#e1f5ee", lineHeight: 1.2 }}>{sp.name || "ชื่อร้าน"}</div>
                <div style={{ fontSize: 10, color: "#9fe1cb" }}>{sp.nameEn || "คำบรรยาย"}</div>
              </div>
            </div>
            <div style={{ fontSize: 11, color: "#9ca3af", textAlign: "center", marginTop: 6 }}>ตัวอย่าง sidebar</div>
          </div>

          <div style={{ flex: 1, minWidth: 200 }}>
            <Field label="ชื่อร้าน (บรรทัดบนใน sidebar)">
              <input style={inputStyle} value={sp.name || ""} onChange={(e) => setSP("name", e.target.value)} placeholder="เช่น วงจรกรีน" />
            </Field>
            <Field label="คำบรรยาย (บรรทัดล่างใน sidebar)">
              <input style={inputStyle} value={sp.nameEn || ""} onChange={(e) => setSP("nameEn", e.target.value)} placeholder="เช่น ระบบซื้อขายของเก่ารีไซเคิล" />
            </Field>
            <Field label="โลโก้ sidebar">
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <label style={{ ...btnSecondary, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6, fontSize: 13 }}>
                  <Image size={14} /> อัปโหลดโลโก้ sidebar
                  <input type="file" accept="image/*" onChange={handleSidebarLogoUpload} style={{ display: "none" }} />
                </label>
                {sp.logo && (
                  <button style={btnDanger} onClick={() => setSP("logo", "")}>
                    <X size={14} /> ลบ
                  </button>
                )}
              </div>
            </Field>
          </div>
        </div>
      </div>

      {/* ===== ส่วนที่ 2: โลโก้บิล/เอกสาร ===== */}
      <div style={sCard}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
          <div style={{ width: 28, height: 28, borderRadius: 6, background: "#185fa5", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <FileText size={14} color="#fff" />
          </div>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#185fa5" }}>ข้อมูลเอกสาร / บิล</h3>
          <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 4 }}>แสดงในใบรับสินค้า ใบขาย และเอกสารทุกใบ</span>
        </div>

        <h4 style={{ margin: "0 0 12px", fontSize: 13, fontWeight: 600, color: "#6b7280" }}>โลโก้บนบิล</h4>
        <div style={{ display: "flex", gap: 20, alignItems: "flex-start", marginBottom: 20, flexWrap: "wrap" }}>
          <div style={{ flexShrink: 0 }}>
            {cs.logo ? (
              <div style={{ position: "relative" }}>
                <img src={cs.logo} alt="โลโก้บิล" style={{ width: 140, height: 88, objectFit: "contain", borderRadius: 10, border: "1px solid #e5e7eb", background: "#f9fafb", padding: 8 }} />
                <button onClick={() => set("logo", "")} style={{ position: "absolute", top: -8, right: -8, background: "#ef4444", border: "none", borderRadius: "50%", width: 22, height: 22, cursor: "pointer", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <X size={12} />
                </button>
              </div>
            ) : (
              <div style={{ width: 140, height: 88, borderRadius: 10, border: "2px dashed #d1d5db", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 4, color: "#9ca3af", background: "#f9fafb" }}>
                <FileText size={24} />
                <span style={{ fontSize: 11 }}>ยังไม่มีโลโก้บิล</span>
              </div>
            )}
          </div>
          <div>
            <label style={{ ...btnSecondary, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6, fontSize: 13, marginBottom: 8 }}>
              <Image size={14} /> อัปโหลดโลโก้บิล
              <input type="file" accept="image/*" onChange={handleLogoUpload} style={{ display: "none" }} />
            </label>
            <p style={{ fontSize: 12, color: "#9ca3af", margin: 0 }}>JPG, PNG, SVG — ไม่เกิน 2MB</p>
          </div>
        </div>

        <h4 style={{ margin: "0 0 12px", fontSize: 13, fontWeight: 600, color: "#6b7280" }}>ข้อมูลร้านบนบิล</h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
          <Field label="ชื่อร้าน / บริษัท (ภาษาไทย)">
            <input style={inputStyle} value={cs.name || ""} onChange={(e) => set("name", e.target.value)} placeholder="เช่น วงจรกรีน รีไซเคิล" />
          </Field>
          <Field label="ชื่อร้าน / บริษัท (English)">
            <input style={inputStyle} value={cs.nameEn || ""} onChange={(e) => set("nameEn", e.target.value)} />
          </Field>
          <Field label="เลขประจำตัวผู้เสียภาษี">
            <input style={inputStyle} value={cs.taxId || ""} onChange={(e) => set("taxId", e.target.value)} placeholder="0-0000-00000-00-0" />
          </Field>
          <Field label="เบอร์โทรศัพท์">
            <input style={inputStyle} value={cs.phone || ""} onChange={(e) => set("phone", e.target.value)} />
          </Field>
          <Field label="อีเมล">
            <input style={inputStyle} value={cs.email || ""} onChange={(e) => set("email", e.target.value)} />
          </Field>
          <Field label="เว็บไซต์ / Line ID">
            <input style={inputStyle} value={cs.website || ""} onChange={(e) => set("website", e.target.value)} />
          </Field>
        </div>
        <Field label="ที่อยู่">
          <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={cs.address || ""} onChange={(e) => set("address", e.target.value)} />
        </Field>

        <h4 style={{ margin: "16px 0 12px", fontSize: 13, fontWeight: 600, color: "#6b7280" }}>ตั้งค่าเอกสาร</h4>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
          <Field label="ชื่อเอกสาร ใบรับสินค้า">
            <input style={inputStyle} value={cs.purchaseTitle || ""} onChange={(e) => set("purchaseTitle", e.target.value)} placeholder="ใบรับสินค้า (รับซื้อของเก่า)" />
          </Field>
          <Field label="ชื่อเอกสาร ใบขาย">
            <input style={inputStyle} value={cs.salesTitle || ""} onChange={(e) => set("salesTitle", e.target.value)} placeholder="ใบกำกับภาษี / Invoice" />
          </Field>
          <Field label="ชื่อเอกสาร ใบสำคัญจ่าย">
            <input style={inputStyle} value={cs.expenseVoucherTitle || ""} onChange={(e) => set("expenseVoucherTitle", e.target.value)} placeholder="ใบสำคัญจ่าย" />
          </Field>
          <div></div>
          <Field label="สีหลักเอกสาร">
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input type="color" value={cs.primaryColor || "#0f6e56"} onChange={(e) => set("primaryColor", e.target.value)} style={{ width: 40, height: 36, border: "1px solid #e5e7eb", borderRadius: 6, cursor: "pointer" }} />
              <input style={{ ...inputStyle, flex: 1 }} value={cs.primaryColor || "#0f6e56"} onChange={(e) => set("primaryColor", e.target.value)} />
            </div>
          </Field>
          <Field label="สีรองเอกสาร">
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input type="color" value={cs.accentColor || "#185fa5"} onChange={(e) => set("accentColor", e.target.value)} style={{ width: 40, height: 36, border: "1px solid #e5e7eb", borderRadius: 6, cursor: "pointer" }} />
              <input style={{ ...inputStyle, flex: 1 }} value={cs.accentColor || "#185fa5"} onChange={(e) => set("accentColor", e.target.value)} />
            </div>
          </Field>
        </div>
        <Field label="หมายเหตุท้ายใบสำคัญจ่าย">
          <input style={inputStyle} value={cs.expenseVoucherNote || ""} onChange={(e) => set("expenseVoucherNote", e.target.value)} placeholder="เช่น ผู้จ่ายเงิน _________________________ ผู้อนุมัติ _________________________" />
        </Field>
        <Field label="หมายเหตุท้ายเอกสาร">
          <textarea style={{ ...inputStyle, minHeight: 56, resize: "vertical" }} value={cs.footerNote || ""} onChange={(e) => set("footerNote", e.target.value)} placeholder="เช่น ขอบคุณที่ใช้บริการ" />
        </Field>
      </div>

      {/* ===== Preview บิล ===== */}
      <div style={{ ...sCard, background: "#f9fafb" }}>
        <h3 style={{ margin: "0 0 12px", fontSize: 14, fontWeight: 700 }}>👁 ตัวอย่างหัวบิล</h3>
        <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 10, padding: 18 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${cs.primaryColor || "#0f6e56"}`, paddingBottom: 10, marginBottom: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {cs.logo ? (
                <img src={cs.logo} alt="logo" style={{ height: 48, maxWidth: 90, objectFit: "contain" }} />
              ) : (
                <div style={{ width: 48, height: 48, background: "#f3f4f6", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" }}>
                  <FileText size={22} color="#9ca3af" />
                </div>
              )}
              <div>
                <div style={{ fontSize: 15, fontWeight: 700, color: cs.primaryColor || "#0f6e56" }}>{cs.name || "ชื่อร้านบนบิล"}</div>
                {cs.taxId && <div style={{ fontSize: 11, color: "#6b7280" }}>เลขผู้เสียภาษี: {cs.taxId}</div>}
                {cs.phone && <div style={{ fontSize: 11, color: "#6b7280" }}>โทร: {cs.phone}</div>}
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: cs.primaryColor || "#0f6e56" }}>{cs.purchaseTitle || "ใบรับสินค้า"}</div>
              <div style={{ fontSize: 11, color: "#6b7280" }}>เลขที่: PO260617001</div>
            </div>
          </div>
          {cs.footerNote && <div style={{ fontSize: 11, color: "#6b7280" }}>{cs.footerNote}</div>}
        </div>
      </div>
    </div>
  );
}

// ===================================================================
// TAX SUMMARY TAB (สรุปภาษีซื้อ-ภาษีขาย)
// ===================================================================
function TaxSummaryTab({ purchases, sales, expenses }) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const [mode, setMode] = useState("month"); // "month" | "range"
  const [rangeStart, setRangeStart] = useState(now.toISOString().slice(0, 8) + "01");
  const [rangeEnd, setRangeEnd] = useState(now.toISOString().slice(0, 10));

  const MONTH_NAMES = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  const yearOptions = [];
  for (let y = 2024; y <= now.getFullYear() + 2; y++) yearOptions.push(y);

  const startDate = mode === "month" ? `${year}-${String(month).padStart(2,"0")}-01` : rangeStart;
  const endDate   = mode === "month" ? `${year}-${String(month).padStart(2,"0")}-${String(new Date(year, month, 0).getDate()).padStart(2,"0")}` : rangeEnd;
  const inRange = (d) => d >= startDate && d <= endDate;
  const periodLabel = mode === "month" ? `${MONTH_NAMES[month]} ${year}` : `${startDate} ถึง ${endDate}`;

  // ===== ภาษีซื้อ (Input VAT) = VAT จากการซื้อสินค้า + ค่าใช้จ่าย =====
  const purchaseVatRows = purchases.filter((po) => po.status === "อนุมัติแล้ว" && inRange(po.date) && Number(po.vatRate) > 0).map((po) => {
    const subtotal = po.items.reduce((s, it) => s + (it.net || 0) * (it.price || 0), 0);
    const vat = subtotal * ((Number(po.vatRate) || 0) / 100);
    return { id: po.id, date: po.date, description: `ใบรับสินค้า ${po.id}`, base: subtotal, vatRate: po.vatRate, vat };
  });

  const expenseVatRows = expenses.filter((e) => inRange(e.billDate || e.date)).flatMap((e) => {
    const items = (e.items && e.items.length > 0) ? e.items : [{ amount: e.amount, vatEnabled: e.vatEnabled, description: e.description }];
    return items.filter((it) => it.vatEnabled).map((it, i) => {
      const base = Number(it.amount) || 0;
      const vat = base * 0.07;
      return { id: `${e.refNo || e.id}${items.length > 1 ? `-${i + 1}` : ""}`, date: e.billDate || e.date, description: `ค่าใช้จ่าย ${e.refNo || e.id}${it.description ? ` (${it.description})` : ""}`, base, vatRate: 7, vat };
    });
  });

  const inputVatRows = [...purchaseVatRows, ...expenseVatRows].sort((a, b) => a.date.localeCompare(b.date));
  const totalInputBase = inputVatRows.reduce((s, r) => s + r.base, 0);
  const totalInputVat  = inputVatRows.reduce((s, r) => s + r.vat, 0);

  // ===== ภาษีขาย (Output VAT) = VAT จากการขายสินค้า =====
  const outputVatRows = sales.filter((inv) => inRange(inv.date) && Number(inv.vatRate) > 0).map((inv) => {
    const subtotal = inv.items.reduce((s, it) => s + (it.net || 0) * (it.price || 0), 0);
    const ad = subtotal - (inv.discount || 0);
    const vat = ad * ((Number(inv.vatRate) || 0) / 100);
    return { id: inv.id, date: inv.date, description: `ใบขาย ${inv.id}`, base: ad, vatRate: inv.vatRate, vat };
  }).sort((a, b) => a.date.localeCompare(b.date));

  const totalOutputBase = outputVatRows.reduce((s, r) => s + r.base, 0);
  const totalOutputVat  = outputVatRows.reduce((s, r) => s + r.vat, 0);
  const vatDiff = totalOutputVat - totalInputVat;
  const vatTh = { ...thStyle, textAlign: "right" };

  // ===== หัก ณ ที่จ่าย (Withholding Tax) =====
  const whtRows = expenses.filter((e) => inRange(e.billDate || e.date)).flatMap((e) => {
    const items = (e.items && e.items.length > 0) ? e.items : [{ amount: e.amount, whtRate: e.whtRate, description: e.description }];
    const vendor = e.vendorName || e.description || e.refNo || e.id;
    return items.filter((it) => Number(it.whtRate) > 0).map((it, i) => {
      const base = Number(it.amount) || 0;
      const wht = base * ((Number(it.whtRate) || 0) / 100);
      return { id: `${e.refNo || e.id}${items.length > 1 ? `-${i + 1}` : ""}`, date: e.billDate || e.date, description: it.description ? `${vendor} (${it.description})` : vendor, base, whtRate: it.whtRate, wht };
    });
  }).sort((a, b) => a.date.localeCompare(b.date));
  const totalWhtBase = whtRows.reduce((s, r) => s + r.base, 0);
  const totalWht     = whtRows.reduce((s, r) => s + r.wht, 0);

  return (
    <div>
      <Header title="สรุปภาษีซื้อ - ภาษีขาย" subtitle="สรุป VAT จากใบรับสินค้า ค่าใช้จ่าย และใบขาย เพื่อยื่น ภ.พ.30">
        <ExportToolbar
          onPDF={() => printAsPDF("tax-content", `ภาษี ${periodLabel}`)}
          onExcel={() => {
            const rows = [
              [`สรุปภาษีซื้อ-ภาษีขาย ${periodLabel}`],[""],
              ["ภาษีซื้อ (Input VAT)","","",""],
              ["เลขที่","วันที่","ฐานภาษี","VAT"],
              ...inputVatRows.map(r => [r.id, r.date, r.base, r.vat]),
              ["รวมภาษีซื้อ","",totalInputBase,totalInputVat],[""],
              ["ภาษีขาย (Output VAT)","","",""],
              ["เลขที่","วันที่","ฐานภาษี","VAT"],
              ...outputVatRows.map(r => [r.id, r.date, r.base, r.vat]),
              ["รวมภาษีขาย","",totalOutputBase,totalOutputVat],[""],
              ["หัก ณ ที่จ่าย (WHT)","","",""],
              ["เลขที่","วันที่","ฐานภาษี","WHT"],
              ...whtRows.map(r => [r.id, r.date, r.base, r.wht]),
              ["รวม WHT","",totalWhtBase,totalWht],[""],
              ["ภาษีสุทธิ (ขาย-ซื้อ)","","",vatDiff],
            ];
            exportExcel(rows, `ภาษี_${periodLabel.replace(/\s/g,"_")}.xlsx`, "ภาษี");
          }}
          onImage={() => printAsPDF("tax-content", `ภาษี ${periodLabel}`)}
        />
      </Header>

      {/* Period Selector */}
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ display: "flex", borderRadius: 8, overflow: "hidden", border: "1px solid #d1d5db" }}>
          {[{key:"month",label:"รายเดือน"},{key:"range",label:"เลือกช่วง"}].map((opt) => (
            <button key={opt.key} onClick={() => setMode(opt.key)}
              style={{ padding:"7px 14px", border:"none", cursor:"pointer", fontSize:13, fontWeight:600,
                background: mode===opt.key ? "#0c443c" : "#fff", color: mode===opt.key ? "#fff" : "#6b7280" }}>
              {opt.label}
            </button>
          ))}
        </div>
        {mode === "month" && (
          <>
            <select style={{ ...inputStyle, width: 140 }} value={month} onChange={(e) => setMonth(Number(e.target.value))}>
              {MONTH_NAMES.slice(1).map((n, i) => <option key={i+1} value={i+1}>{n}</option>)}
            </select>
            <select style={{ ...inputStyle, width: 100 }} value={year} onChange={(e) => setYear(Number(e.target.value))}>
              {yearOptions.map((y) => <option key={y} value={y}>ปี {y}</option>)}
            </select>
          </>
        )}
        {mode === "range" && (
          <>
            <input type="date" style={{ ...inputStyle, width: 160 }} value={rangeStart} onChange={(e) => setRangeStart(e.target.value)} />
            <span style={{ fontSize: 13, color: "#6b7280" }}>ถึง</span>
            <input type="date" style={{ ...inputStyle, width: 160 }} value={rangeEnd} onChange={(e) => setRangeEnd(e.target.value)} />
          </>
        )}
      </div>

      {/* Summary Cards */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 20 }}>
        <div style={{ background: "#faece7", borderRadius: 12, padding: "14px 18px" }}>
          <div style={{ fontSize: 12, color: "#993c1d", marginBottom: 4 }}>ภาษีซื้อ (Input VAT)</div>
          <div style={{ fontWeight: 700, fontSize: 20, color: "#993c1d" }}>฿{fmt(totalInputVat)}</div>
          <div style={{ fontSize: 11, color: "#9ca3af" }}>ฐานภาษี ฿{fmt(totalInputBase)}</div>
        </div>
        <div style={{ background: "#e1f5ee", borderRadius: 12, padding: "14px 18px" }}>
          <div style={{ fontSize: 12, color: "#0f6e56", marginBottom: 4 }}>ภาษีขาย (Output VAT)</div>
          <div style={{ fontWeight: 700, fontSize: 20, color: "#0f6e56" }}>฿{fmt(totalOutputVat)}</div>
          <div style={{ fontSize: 11, color: "#9ca3af" }}>ฐานภาษี ฿{fmt(totalOutputBase)}</div>
        </div>
        <div style={{ background: vatDiff >= 0 ? "#e6f1fb" : "#faeeda", borderRadius: 12, padding: "14px 18px" }}>
          <div style={{ fontSize: 12, color: vatDiff >= 0 ? "#185fa5" : "#854f0b", marginBottom: 4 }}>
            {vatDiff >= 0 ? "VAT ต้องชำระ" : "VAT ขอคืน"}
          </div>
          <div style={{ fontWeight: 700, fontSize: 20, color: vatDiff >= 0 ? "#185fa5" : "#854f0b" }}>฿{fmt(Math.abs(vatDiff))}</div>
        </div>
        <div style={{ background: "#eeedfe", borderRadius: 12, padding: "14px 18px" }}>
          <div style={{ fontSize: 12, color: "#3c3489", marginBottom: 4 }}>หัก ณ ที่จ่าย (WHT)</div>
          <div style={{ fontWeight: 700, fontSize: 20, color: "#3c3489" }}>฿{fmt(totalWht)}</div>
          <div style={{ fontSize: 11, color: "#9ca3af" }}>ฐานภาษี ฿{fmt(totalWhtBase)}</div>
        </div>
      </div>

      <div id="tax-content">
        {/* ภาษีซื้อ */}
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "hidden", marginBottom: 16 }}>
          <div style={{ background: "#faece7", padding: "10px 16px", fontWeight: 700, fontSize: 14, color: "#993c1d" }}>
            ภาษีซื้อ (Input VAT) — {periodLabel}
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={thStyle}>เลขที่เอกสาร</th><th style={thStyle}>วันที่</th>
              <th style={thStyle}>รายการ</th><th style={vatTh}>อัตรา VAT</th>
              <th style={vatTh}>ฐานภาษี (บาท)</th><th style={vatTh}>VAT (บาท)</th>
            </tr></thead>
            <tbody>
              {inputVatRows.map((r,i) => (
                <tr key={i}>
                  <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{r.id}</td>
                  <td style={tdStyle}>{r.date}</td>
                  <td style={tdStyle}>{r.description}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{r.vatRate}%</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(r.base)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#993c1d" }}>{fmt(r.vat)}</td>
                </tr>
              ))}
              {inputVatRows.length === 0 && <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีรายการภาษีซื้อในช่วงนี้</td></tr>}
            </tbody>
            {inputVatRows.length > 0 && <tfoot>
              <tr>
                <td colSpan={4} style={{ ...tdStyle, fontWeight: 700 }}>รวมภาษีซื้อ</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(totalInputBase)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>{fmt(totalInputVat)}</td>
              </tr>
            </tfoot>}
          </table>
        </div>

        {/* ภาษีขาย */}
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "hidden", marginBottom: 16 }}>
          <div style={{ background: "#e1f5ee", padding: "10px 16px", fontWeight: 700, fontSize: 14, color: "#0f6e56" }}>
            ภาษีขาย (Output VAT) — {periodLabel}
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={thStyle}>เลขที่เอกสาร</th><th style={thStyle}>วันที่</th>
              <th style={thStyle}>รายการ</th><th style={vatTh}>อัตรา VAT</th>
              <th style={vatTh}>ฐานภาษี (บาท)</th><th style={vatTh}>VAT (บาท)</th>
            </tr></thead>
            <tbody>
              {outputVatRows.map((r,i) => (
                <tr key={i}>
                  <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{r.id}</td>
                  <td style={tdStyle}>{r.date}</td>
                  <td style={tdStyle}>{r.description}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{r.vatRate}%</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(r.base)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#0f6e56" }}>{fmt(r.vat)}</td>
                </tr>
              ))}
              {outputVatRows.length === 0 && <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีรายการภาษีขายในช่วงนี้</td></tr>}
            </tbody>
            {outputVatRows.length > 0 && <tfoot>
              <tr>
                <td colSpan={4} style={{ ...tdStyle, fontWeight: 700 }}>รวมภาษีขาย</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(totalOutputBase)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>{fmt(totalOutputVat)}</td>
              </tr>
            </tfoot>}
          </table>
        </div>

        {/* สรุป */}
        <div style={{ background: "#fff", borderRadius: 12, border: "2px solid #0c443c", padding: "18px 20px" }}>
          <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 12, color: "#0c443c" }}>สรุปภาษีสุทธิ — {periodLabel}</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, fontSize: 14 }}>
            <div><span style={{ color: "#6b7280" }}>ภาษีขาย</span><div style={{ fontWeight: 700, fontSize: 18, color: "#0f6e56" }}>฿{fmt(totalOutputVat)}</div></div>
            <div><span style={{ color: "#6b7280" }}>หัก ภาษีซื้อ</span><div style={{ fontWeight: 700, fontSize: 18, color: "#993c1d" }}>฿{fmt(totalInputVat)}</div></div>
            <div>
              <span style={{ color: "#6b7280" }}>{vatDiff >= 0 ? "ภาษีต้องชำระ" : "ภาษีขอคืน"}</span>
              <div style={{ fontWeight: 700, fontSize: 20, color: vatDiff >= 0 ? "#185fa5" : "#854f0b" }}>฿{fmt(Math.abs(vatDiff))}</div>
              <div style={{ fontSize: 11, color: "#9ca3af" }}>{vatDiff >= 0 ? "นำส่งกรมสรรพากร" : "ยื่นขอคืนภาษี"}</div>
            </div>
          </div>
        </div>

        {/* หัก ณ ที่จ่าย */}
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", overflow: "hidden", marginTop: 16 }}>
          <div style={{ background: "#eeedfe", padding: "10px 16px", fontWeight: 700, fontSize: 14, color: "#3c3489" }}>
            หัก ณ ที่จ่าย (Withholding Tax) — {periodLabel}
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={thStyle}>เลขที่เอกสาร</th>
              <th style={thStyle}>วันที่</th>
              <th style={thStyle}>ผู้รับเงิน / รายการ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>อัตรา WHT</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ฐานภาษี (บาท)</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ภาษีหัก ณ ที่จ่าย (บาท)</th>
            </tr></thead>
            <tbody>
              {whtRows.map((r, i) => (
                <tr key={i}>
                  <td style={{ ...tdStyle, fontFamily: "monospace", fontSize: 12 }}>{r.id}</td>
                  <td style={tdStyle}>{r.date}</td>
                  <td style={tdStyle}>{r.description}</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{r.whtRate}%</td>
                  <td style={{ ...tdStyle, textAlign: "right" }}>{fmt(r.base)}</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#3c3489" }}>{fmt(r.wht)}</td>
                </tr>
              ))}
              {whtRows.length === 0 && <tr><td colSpan={6} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ไม่มีรายการหัก ณ ที่จ่ายในช่วงนี้</td></tr>}
            </tbody>
            {whtRows.length > 0 && <tfoot><tr>
              <td colSpan={4} style={{ ...tdStyle, fontWeight: 700 }}>รวมภาษีหัก ณ ที่จ่าย</td>
              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(totalWhtBase)}</td>
              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#3c3489" }}>{fmt(totalWht)}</td>
            </tr></tfoot>}
          </table>
        </div>
      </div>
    </div>
  );
}

// ===================================================================
// DELIVERY TAB (ใบส่งสินค้า)
// ===================================================================
function DeliveryTab({ deliveries, setDeliveries, customers, sales, products, companySettings }) {
  const cs = companySettings || {};
  const [modal, setModal] = useState(null);
  const [search, setSearch] = useState("");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const custName = (id) => customers.find((c) => c.id === id)?.name || id;
  const prodName = (id) => products.find((p) => p.id === id)?.name || id;
  const prodUnit = (id) => products.find((p) => p.id === id)?.unit || "";

  // ประเภทภาชนะที่เคยพิมพ์ไว้ — สะสมเป็นตัวเลือกให้พิมพ์ซ้ำง่ายขึ้น
  const containerTypeOptions = [...new Set(deliveries.flatMap((d) => (d.items || []).map((it) => it.containerType)).filter(Boolean))];

  const blankItem = () => ({ id: "DI" + Date.now().toString().slice(-6) + Math.floor(Math.random() * 1000), productId: "", qty: 0, containerWeight: 0, containerType: "" });
  const blankForm = () => ({
    id: genId("DV", deliveries, new Date().toISOString().slice(0, 10)),
    date: new Date().toISOString().slice(0, 10),
    customerId: "",
    relatedSaleId: "",
    items: [blankItem()],
    vehicleNo: "",
    driverName: "",
    note: "",
  });
  const [form, setForm] = useState(blankForm());

  const openAdd = () => { setForm(blankForm()); setModal({ mode: "add" }); };
  const openEdit = (item) => { setForm(JSON.parse(JSON.stringify(item))); setModal({ mode: "edit", item }); };
  const openView = (item) => setModal({ mode: "view", item });

  const addItem = () => setForm({ ...form, items: [...form.items, blankItem()] });
  const updateItem = (idx, field, value) => {
    const items = [...form.items];
    items[idx] = { ...items[idx], [field]: value };
    setForm({ ...form, items });
  };
  const removeItem = (idx) => setForm({ ...form, items: form.items.filter((_, i) => i !== idx) });
  const netQtyOf = (it) => (Number(it.qty) || 0) - (Number(it.containerWeight) || 0);

  const save = () => {
    if (!form.customerId || form.items.length === 0) return;
    const cleaned = { ...form, items: form.items.map((it) => ({ ...it, qty: Number(it.qty) || 0, containerWeight: Number(it.containerWeight) || 0 })) };
    if (modal.mode === "add") setDeliveries([...deliveries, cleaned]);
    else setDeliveries(deliveries.map((d) => (d.id === modal.item.id ? cleaned : d)));
    setModal(null);
  };

  const remove = (id) => setDeliveries(deliveries.filter((d) => d.id !== id));

  const filtered = deliveries
    .filter((d) => d.id.includes(search) || custName(d.customerId).includes(search) || (d.vehicleNo || "").includes(search))
    .filter((d) => (!dateFrom || (d.date || "") >= dateFrom) && (!dateTo || (d.date || "") <= dateTo))
    .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));

  const deliveryNetTotal = (d) => (d.items || []).reduce((s, it) => s + netQtyOf(it), 0);

  const customerSales = (custId) => sales.filter((s) => s.customerId === custId);

  return (
    <div>
      <Header title="ใบส่งสินค้า" subtitle="บันทึกและพิมพ์ใบส่งสินค้าให้ลูกค้า">
        <div style={{ display: "flex", gap: 8 }}>
          <ExportToolbar
            onPDF={() => printAsPDF("tab-export-delivery", "ใบส่งสินค้า")}
            onExcel={() => {
              const rows = [
                ["เลขที่", "วันที่", "ลูกค้า", "ทะเบียนรถ", "คนขับ", "ยอดรวมสุทธิ"],
                ...filtered.map((d) => [d.id, d.date, custName(d.customerId), d.vehicleNo || "", d.driverName || "", deliveryNetTotal(d)]),
              ];
              exportExcel(rows, "ใบส่งสินค้า.xlsx", "ใบส่งสินค้า");
            }}
            onImage={() => printAsPDF("tab-export-delivery", "ใบส่งสินค้า")}
          />
          <button style={btnPrimary} onClick={openAdd}><Plus size={16} /> สร้างใบส่งสินค้า</button>
        </div>
      </Header>

      <div id="tab-export-delivery">
      <SearchBar value={search} onChange={setSearch} placeholder="ค้นหาเลขที่, ลูกค้า, ทะเบียนรถ..." dateFrom={dateFrom} dateTo={dateTo} onDateFromChange={setDateFrom} onDateToChange={setDateTo} />

      <Card>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              <th style={thStyle}>เลขที่</th>
              <th style={thStyle}>วันที่</th>
              <th style={thStyle}>ลูกค้า</th>
              <th style={thStyle}>ทะเบียนรถ</th>
              <th style={thStyle}>คนขับ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ยอดรวมสุทธิ</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((d) => (
              <tr key={d.id}>
                <td style={{ ...tdStyle, fontFamily: "'JetBrains Mono', monospace", color: "#534ab7" }}>{d.id}</td>
                <td style={tdStyle}>{d.date}</td>
                <td style={tdStyle}>{custName(d.customerId)}</td>
                <td style={tdStyle}>{d.vehicleNo || "-"}</td>
                <td style={tdStyle}>{d.driverName || "-"}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>{fmt(deliveryNetTotal(d))}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>
                  <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                    <button style={iconBtn} onClick={() => openView(d)}><Printer size={14} /> พิมพ์</button>
                    <button style={iconBtn} onClick={() => openEdit(d)}><Edit2 size={14} /> แก้ไข</button>
                    <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบใบส่งสินค้า "${d.id}" ใช่หรือไม่?`, () => remove(d.id))}><Trash2 size={14} /> ลบ</button>
                  </div>
                </td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={7} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีใบส่งสินค้า</td></tr>}
          </tbody>
        </table>
      </Card>
      </div>

      {modal && (modal.mode === "add" || modal.mode === "edit") && (
        <Modal title={modal.mode === "add" ? "สร้างใบส่งสินค้า" : "แก้ไขใบส่งสินค้า"} onClose={() => setModal(null)} wide>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            <Field label="เลขที่ใบส่ง"><input style={inputStyle} value={form.id} onChange={(e) => setForm({ ...form, id: e.target.value })} placeholder="เช่น DV260622001" /></Field>
            <Field label="วันที่"><input type="date" style={inputStyle} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value, id: genId("DV", deliveries, e.target.value) })} /></Field>
            <Field label="ลูกค้า">
              <CustomerSelect customers={customers} value={form.customerId} onChange={(cid) => setForm({ ...form, customerId: cid, relatedSaleId: "" })} labelWithId={false} />
            </Field>
          </div>
          {customerSales(form.customerId).length > 0 && (
            <Field label="อ้างอิงใบขาย (ถ้ามี)">
              <select style={inputStyle} value={form.relatedSaleId} onChange={(e) => setForm({ ...form, relatedSaleId: e.target.value })}>
                <option value="">-- ไม่อ้างอิงใบขาย --</option>
                {customerSales(form.customerId).map((s) => <option key={s.id} value={s.id}>{s.id} ({s.date})</option>)}
              </select>
            </Field>
          )}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 16px" }}>
            <Field label="ทะเบียนรถ"><input style={inputStyle} value={form.vehicleNo} onChange={(e) => setForm({ ...form, vehicleNo: e.target.value })} /></Field>
            <Field label="ชื่อคนขับ"><input style={inputStyle} value={form.driverName} onChange={(e) => setForm({ ...form, driverName: e.target.value })} /></Field>
          </div>

          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <div style={{ fontWeight: 600, fontSize: 14 }}>รายการสินค้า</div>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 8 }}>
            <thead>
              <tr>
                <th style={thStyle}>สินค้า</th>
                <th style={{ ...thStyle, textAlign: "right" }}>จำนวน</th>
                <th style={{ ...thStyle, textAlign: "right" }}>น้ำหนักภาชนะ</th>
                <th style={thStyle}>ประเภทภาชนะ</th>
                <th style={{ ...thStyle, textAlign: "right" }}>จำนวนสุทธิ</th>
                <th style={thStyle}></th>
              </tr>
            </thead>
            <tbody>
              {form.items.map((it, idx) => (
                <tr key={it.id}>
                  <td style={tdStyle}><ProductSelect products={products} value={it.productId} onChange={(pid) => updateItem(idx, "productId", pid)} /></td>
                  <td style={tdStyle}><input type="number" style={{ ...inputStyle, width: 90, textAlign: "right" }} value={it.qty} onChange={(e) => updateItem(idx, "qty", e.target.value)} /></td>
                  <td style={tdStyle}><input type="number" style={{ ...inputStyle, width: 90, textAlign: "right" }} value={it.containerWeight} onChange={(e) => updateItem(idx, "containerWeight", e.target.value)} /></td>
                  <td style={tdStyle}>
                    <input style={{ ...inputStyle, width: 110 }} list="delivery-container-type-options" value={it.containerType} onChange={(e) => updateItem(idx, "containerType", e.target.value)} placeholder="เช่น ถุง, ลัง" />
                  </td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600 }}>{fmt(netQtyOf(it))} {prodUnit(it.productId)}</td>
                  <td style={tdStyle}><button style={btnDanger} onClick={() => removeItem(idx)}><Trash2 size={14} /></button></td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(form.items.reduce((s, it) => s + (Number(it.qty) || 0), 0))} กก.</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>{fmt(form.items.reduce((s, it) => s + (Number(it.containerWeight) || 0), 0))} กก.</td>
                <td style={tdStyle}></td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>{fmt(form.items.reduce((s, it) => s + netQtyOf(it), 0))} กก.</td>
                <td style={tdStyle}></td>
              </tr>
            </tfoot>
          </table>
          <datalist id="delivery-container-type-options">
            {containerTypeOptions.map((c) => <option key={c} value={c} />)}
          </datalist>
          <button style={btnSecondary} onClick={addItem}><Plus size={14} /> เพิ่มรายการ</button>

          <Field label="หมายเหตุ"><input style={inputStyle} value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ยกเลิก</button>
            <button style={btnPrimary} onClick={save}><Save size={16} /> บันทึก</button>
          </div>
        </Modal>
      )}

      {modal && modal.mode === "view" && (
        <Modal title={`ใบส่งสินค้า · ${modal.item.id}`} onClose={() => setModal(null)} wide>
          <div id="delivery-pdf-content" style={{ background: "#fff", padding: 24, border: "1px solid #e5e7eb", borderRadius: 8 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${cs.accentColor || "#185fa5"}`, paddingBottom: 12, marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                {cs.logo && (
                  <img src={cs.logo} alt="logo" style={{ height: 50, maxWidth: 100, objectFit: "contain" }} />
                )}
                <div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: cs.accentColor || "#185fa5" }}>{cs.name || "วงจรกรีน รีไซเคิล"}</div>
                  {cs.taxId && <div style={{ fontSize: 12, color: "#6b7280" }}>เลขผู้เสียภาษี: {cs.taxId}</div>}
                  {cs.address && <div style={{ fontSize: 12, color: "#6b7280" }}>{cs.address}</div>}
                  {cs.phone && <div style={{ fontSize: 12, color: "#6b7280" }}>โทร: {cs.phone}</div>}
                  <div style={{ fontSize: 12, color: "#6b7280" }}>ใบส่งสินค้า</div>
                </div>
              </div>
              <div style={{ textAlign: "right", fontSize: 12, color: "#6b7280" }}>
                <div>เลขที่: {modal.item.id}</div>
                <div>วันที่: {modal.item.date}</div>
                {modal.item.relatedSaleId && <div>อ้างอิงใบขาย: {modal.item.relatedSaleId}</div>}
              </div>
            </div>
            <div style={{ marginBottom: 12, fontSize: 13 }}>
              <div><strong>ลูกค้า:</strong> {custName(modal.item.customerId)}</div>
              {modal.item.vehicleNo && <div><strong>ทะเบียนรถ:</strong> {modal.item.vehicleNo}</div>}
              {modal.item.driverName && <div><strong>คนขับ:</strong> {modal.item.driverName}</div>}
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead><tr style={{ background: "#f3f4f6" }}>
                <th style={{ ...thStyle, padding: "4px 12px" }}>#</th><th style={{ ...thStyle, padding: "4px 12px" }}>สินค้า</th><th style={{ ...thStyle, padding: "4px 12px", textAlign: "right" }}>จำนวน</th><th style={{ ...thStyle, padding: "4px 12px", textAlign: "right" }}>น้ำหนักภาชนะ</th><th style={{ ...thStyle, padding: "4px 12px" }}>ประเภทภาชนะ</th><th style={{ ...thStyle, padding: "4px 12px", textAlign: "right" }}>จำนวนสุทธิ</th>
              </tr></thead>
              <tbody>
                {modal.item.items.map((it, i) => (
                  <tr key={it.id}>
                    <td style={{ ...tdStyle, padding: "4px 12px" }}>{i + 1}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px" }}>{prodName(it.productId)}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right" }}>{fmt(it.qty)}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right" }}>{fmt(it.containerWeight)}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px" }}>{it.containerType || "-"}</td>
                    <td style={{ ...tdStyle, padding: "4px 12px", textAlign: "right", fontWeight: 600 }}>{fmt(netQtyOf(it))} {prodUnit(it.productId)}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{ background: "#f3f4f6", borderTop: "2px solid #e5e7eb" }}>
                  <td colSpan={2} style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งหมด</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>{fmt(modal.item.items.reduce((s, it) => s + (Number(it.qty) || 0), 0))} กก.</td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#993c1d" }}>{fmt(modal.item.items.reduce((s, it) => s + (Number(it.containerWeight) || 0), 0))} กก.</td>
                  <td style={tdStyle}></td>
                  <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>{fmt(deliveryNetTotal(modal.item))} กก.</td>
                </tr>
              </tfoot>
            </table>
            {modal.item.note && <div style={{ marginTop: 12, fontSize: 13, color: "#6b7280" }}>หมายเหตุ: {modal.item.note}</div>}
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 48, fontSize: 12 }}>
              <div style={{ textAlign: "center", width: "45%" }}><div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้ส่งสินค้า</div></div>
              <div style={{ textAlign: "center", width: "45%" }}><div style={{ borderTop: "1px solid #9ca3af", paddingTop: 6 }}>ผู้รับสินค้า</div></div>
            </div>
          </div>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
            <button style={btnSecondary} onClick={() => setModal(null)}>ปิด</button>
            <button style={btnPrimary} onClick={() => printAsPDF("delivery-pdf-content", `ใบส่งสินค้า ${modal.item.id}`)}><Download size={16} /> พิมพ์</button>
          </div>
        </Modal>
      )}
    </div>
  );
}

function Badge({ text }) {
  return <span style={{ display: "inline-block", padding: "2px 10px", borderRadius: 999, fontSize: 11, fontWeight: 600, background: "#f1efe8", color: "#444441" }}>{text}</span>;
}

// ===================================================================
// MONTHLY REPORT TAB (รายงานกำไร/ขาดทุน, สรุปรายเดือน, เงินปันผล)
// ===================================================================
function MonthlyReportTab({ purchases, sales, expenses, deposits, inventory, expenseCategories, shareholders, setShareholders, dividendPayments, setDividendPayments, companySettings, setCompanySettings }) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const [reportView, setReportView] = useState("monthly"); // "monthly" | "yearly"
  const [editingShareholders, setEditingShareholders] = useState(false);
  const openingRevenue = Number(companySettings?.openingRevenue) || 0;
  const openingCost = Number(companySettings?.openingCost) || 0;
  const openingMonth = companySettings?.openingMonth || "";
  const setOpeningRevenue = (v) => setCompanySettings((prev) => ({ ...prev, openingRevenue: Number(v) || 0 }));
  const setOpeningCost = (v) => setCompanySettings((prev) => ({ ...prev, openingCost: Number(v) || 0 }));
  const setOpeningMonth = (v) => setCompanySettings((prev) => ({ ...prev, openingMonth: v }));

  const MONTH_NAMES = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  const yearOptions = [];
  for (let y = 2024; y <= now.getFullYear() + 2; y++) yearOptions.push(y);

  const movements = inventory?.movements || [];
  // มูลค่าสต็อก ณ จุดใดจุดหนึ่ง = ผลรวมมูลค่า "in" ลบมูลค่า "out"/"withdraw" (costConsumed) ของรายการที่เกิดขึ้น "ก่อน" วันที่ที่กำหนด (exclusive)
  const stockValueBefore = (dateExclusive) => {
    let value = 0;
    movements.forEach((m) => {
      if (m.date >= dateExclusive) return;
      if (m.type === "in") value += (Number(m.qty) || 0) * (Number(m.price) || 0);
      else value -= Number(m.costConsumed) || 0;
    });
    return value;
  };

  // ฟังก์ชันคำนวณกำไรขาดทุนของเดือน/ปีใดๆ — ใช้ร่วมกันทั้งมุมมองรายเดือนและสรุปรายปี
  const computeMonthlyPL = (y, m) => {
    const sd = `${y}-${String(m).padStart(2,"0")}-01`;
    const ed = `${y}-${String(m).padStart(2,"0")}-${String(new Date(y, m, 0).getDate()).padStart(2,"0")}`;
    const inR = (d) => d >= sd && d <= ed;
    const ym = `${y}-${String(m).padStart(2,"0")}`;

    const salesInR = sales.filter((s) => inR(s.date));
    const totalRev = salesInR.reduce((sum, inv) => {
      const subtotal = inv.items.reduce((s, it) => s + (it.net || 0) * (it.price || 0), 0);
      const ad = subtotal - (inv.discount || 0);
      return sum + ad;
    }, 0);
    const otherIncome = 0;
    const income = totalRev + otherIncome;

    const beginInv = stockValueBefore(sd);
    const endInv = stockValueBefore(new Date(new Date(ed).getTime() + 86400000).toISOString().slice(0, 10));
    const purchInR = movements
      .filter((mv) => mv.type === "in" && !mv.isOpening && inR(mv.date))
      .reduce((s, mv) => s + (Number(mv.qty) || 0) * (Number(mv.price) || 0), 0);
    const available = beginInv + purchInR;
    const cost = available - endInv;
    const gross = income - cost;

    const expensesInR = expenses.filter((e) => inR(e.billDate || e.date));
    const openingRows = [];
    Object.entries(expenseCategories || {}).forEach(([main, subs]) => {
      (subs || []).forEach((s) => {
        if (typeof s === "string") return;
        if (Number(s.openingBalance) > 0 && s.openingMonth === ym) {
          openingRows.push({ mainCategory: main, amount: Number(s.openingBalance) });
        }
      });
    });
    const groups = {};
    expensesInR.forEach((e) => {
      const items = (e.items && e.items.length > 0) ? e.items : [{ mainCategory: e.mainCategory || e.category || "ไม่ระบุ", amount: e.amount }];
      items.filter((it) => it.mainCategory === "ค่าใช้จ่าย").forEach((it) => {
        groups["ค่าใช้จ่าย"] = (groups["ค่าใช้จ่าย"] || 0) + (Number(it.amount) || 0);
      });
    });
    openingRows.filter((r) => r.mainCategory === "ค่าใช้จ่าย").forEach((r) => { groups["ค่าใช้จ่าย"] = (groups["ค่าใช้จ่าย"] || 0) + r.amount; });
    const byCategory = Object.entries(groups).map(([category, amount]) => ({ category, amount })).sort((a, b) => b.amount - a.amount);
    const totalExp = byCategory.reduce((s, c) => s + c.amount, 0);

    // รวมยอดยกมาเข้าใน function โดยตรง
    const ymMatches = !openingMonth || openingMonth === ym;
    const addRev = ymMatches ? (Number(openingRevenue) || 0) : 0;
    const addCost = ymMatches ? (Number(openingCost) || 0) : 0;
    const totalIncomeWithOpening = income + addRev;
    const totalCostWithOpening = cost + addCost;
    const grossWithOpening = totalIncomeWithOpening - totalCostWithOpening;
    const net = grossWithOpening - totalExp;
    return { totalRevenue: totalRev, totalOtherIncome: otherIncome, totalIncome: totalIncomeWithOpening, beginningInventory: beginInv, endingInventory: endInv, purchasesInRange: purchInR, goodsAvailableForSale: available, totalCost: totalCostWithOpening, grossProfit: grossWithOpening, expenseByCategory: byCategory, totalExpenses: totalExp, netProfit: net, openingRevenueApplied: addRev, openingCostApplied: addCost };
  };

  const startDate = `${year}-${String(month).padStart(2,"0")}-01`;
  const endDate   = `${year}-${String(month).padStart(2,"0")}-${String(new Date(year, month, 0).getDate()).padStart(2,"0")}`;

  const currentMonthPL = computeMonthlyPL(year, month);
  const {
    totalRevenue, totalOtherIncome, totalIncome,
    beginningInventory, endingInventory, purchasesInRange, goodsAvailableForSale, totalCost,
    grossProfit, expenseByCategory, totalExpenses, netProfit,
    openingRevenueApplied, openingCostApplied,
  } = currentMonthPL;
  const openingApplies = openingRevenueApplied > 0 || openingCostApplied > 0;
  const profitMargin = totalIncome > 0 ? (netProfit / totalIncome) * 100 : 0;

  // ===== สรุปรายปี: กำไรสุทธิทั้ง 12 เดือน =====
  const yearlyMonths = useMemo(() => {
    return Array.from({ length: 12 }, (_, i) => {
      const m = i + 1;
      const pl = computeMonthlyPL(year, m);
      return { month: m, label: MONTH_NAMES[m], ...pl };
    });
  }, [year, sales, expenses, expenseCategories, movements]);
  const yearlyNetProfitTotal = yearlyMonths.reduce((s, m) => s + m.netProfit, 0);

  // ===== บันทึกจ่ายเงินปันผล (รายปี) =====
  const dividendPaymentsThisYear = (dividendPayments || []).filter((d) => (d.date || "").startsWith(String(year)));
  const totalDividendPaidThisYear = dividendPaymentsThisYear.reduce((s, d) => s + (Number(d.amount) || 0), 0);
  const [divPayForm, setDivPayForm] = useState(null); // {date, amount}
  const openDivPayForm = () => setDivPayForm({ id: "DIV" + Date.now().toString().slice(-6), date: new Date().toISOString().slice(0, 10), amount: "" });
  const saveDivPayment = () => {
    if (!divPayForm || !(Number(divPayForm.amount) > 0)) return;
    setDividendPayments([...(dividendPayments || []), { ...divPayForm, amount: Number(divPayForm.amount) }]);
    setDivPayForm(null);
  };
  const removeDivPayment = (id) => setDividendPayments((dividendPayments || []).filter((d) => d.id !== id));


  const totalSharePercent = shareholders.reduce((s, sh) => s + (Number(sh.percent) || 0), 0);
  const dividendPool = Math.max(0, netProfit);

  const updateShareholder = (idx, field, value) => {
    const updated = [...shareholders];
    updated[idx] = { ...updated[idx], [field]: value };
    setShareholders(updated);
  };
  const addShareholder = () => setShareholders([...shareholders, { id: "SH" + (shareholders.length + 1), name: `หุ้นส่วน ${shareholders.length + 1}`, percent: 0 }]);
  const removeShareholder = (idx) => setShareholders(shareholders.filter((_, i) => i !== idx));

  const periodLabel = `${MONTH_NAMES[month]} ${year}`;

  return (
    <div>
      <Header title="รายงานกำไร-ขาดทุนรายเดือน" subtitle="สรุปผลประกอบการและคำนวณเงินปันผลตามสัดส่วนหุ้น">
        <ExportToolbar
          onPDF={() => printAsPDF("monthly-report-content", `รายงาน ${periodLabel}`)}
          onExcel={() => {
            const rows = [
              [`รายงานกำไร-ขาดทุน ${periodLabel}`],[""],
              ["รายได้จากการขาย", totalRevenue],
              ["รายได้อื่น", totalOtherIncome],
              ["รวมรายได้", totalIncome],[""],
              ["สินค้าคงเหลือยกมาต้นงวด", beginningInventory],
              ["บวก ซื้อสินค้า", purchasesInRange],
              ["สินค้าที่มีไว้เพื่อขาย", goodsAvailableForSale],
              ["หัก สินค้าคงเหลือปลายงวด", endingInventory],
              ["ต้นทุนขาย", totalCost],[""],
              ["กำไรขั้นต้น", grossProfit],[""],
              ["ค่าใช้จ่ายดำเนินงาน","" ],
              ...expenseByCategory.map(c => [c.category, c.amount]),
              ["รวมค่าใช้จ่าย", totalExpenses],[""],
              ["กำไรสุทธิ", netProfit],
              ["อัตรากำไรสุทธิ (%)", profitMargin.toFixed(2)],
            ];
            exportExcel(rows, `รายงาน_${periodLabel.replace(/\s/g,"_")}.xlsx`, "รายงาน");
          }}
          onImage={() => printAsPDF("monthly-report-content", `รายงาน ${periodLabel}`)}
        />
      </Header>

      <div style={{ display: "flex", gap: 8, marginBottom: 20, overflowX: "auto", overflowY: "hidden", paddingBottom: 4, WebkitOverflowScrolling: "touch" }}>
        {[
          { key: "monthly", label: "รายเดือน" },
          { key: "yearly", label: "สรุปรายปี" },
        ].map((opt) => (
          <button key={opt.key} onClick={() => setReportView(opt.key)}
            style={{ padding: "8px 18px", borderRadius: 8, cursor: "pointer", fontSize: 13, fontWeight: 600, border: "1px solid",
              borderColor: reportView === opt.key ? "#1d9e75" : "#d1d5db",
              background: reportView === opt.key ? "#e1f5ee" : "#fff",
              color: reportView === opt.key ? "#0f6e56" : "#6b7280" }}>
            {opt.label}
          </button>
        ))}
      </div>

      {reportView === "monthly" && (
      <>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
        <select style={{ ...inputStyle, width: 140 }} value={month} onChange={(e) => setMonth(Number(e.target.value))}>
          {MONTH_NAMES.slice(1).map((n, i) => <option key={i+1} value={i+1}>{n}</option>)}
        </select>
        <select style={{ ...inputStyle, width: 100 }} value={year} onChange={(e) => setYear(Number(e.target.value))}>
          {yearOptions.map((y) => <option key={y} value={y}>ปี {y}</option>)}
        </select>
      </div>

      <div id="monthly-report-content">
        {/* Summary cards */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 20 }}>
          <div style={{ background: "#e1f5ee", borderRadius: 12, padding: "14px 18px" }}>
            <div style={{ fontSize: 12, color: "#0f6e56", marginBottom: 4 }}>รวมรายได้</div>
            <div style={{ fontWeight: 700, fontSize: 20, color: "#0f6e56" }}>฿{fmt(totalIncome)}</div>
          </div>
          <div style={{ background: "#faece7", borderRadius: 12, padding: "14px 18px" }}>
            <div style={{ fontSize: 12, color: "#993c1d", marginBottom: 4 }}>ต้นทุนขาย + ค่าใช้จ่าย</div>
            <div style={{ fontWeight: 700, fontSize: 20, color: "#993c1d" }}>฿{fmt(totalCost + totalExpenses)}</div>
          </div>
          <div style={{ background: netProfit >= 0 ? "#e6f1fb" : "#faeeda", borderRadius: 12, padding: "14px 18px" }}>
            <div style={{ fontSize: 12, color: netProfit >= 0 ? "#185fa5" : "#854f0b", marginBottom: 4 }}>กำไรสุทธิ</div>
            <div style={{ fontWeight: 700, fontSize: 20, color: netProfit >= 0 ? "#185fa5" : "#854f0b" }}>฿{fmt(netProfit)}</div>
          </div>
          <div style={{ background: "#eeedfe", borderRadius: 12, padding: "14px 18px" }}>
            <div style={{ fontSize: 12, color: "#3c3489", marginBottom: 4 }}>อัตรากำไรสุทธิ</div>
            <div style={{ fontWeight: 700, fontSize: 20, color: "#3c3489" }}>{profitMargin.toFixed(1)}%</div>
          </div>
        </div>

        {/* P&L Statement */}
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "20px 24px", marginBottom: 20 }}>
          <h3 style={{ margin: "0 0 16px", fontSize: 15, fontWeight: 700 }}>งบกำไรขาดทุน — {periodLabel}</h3>

          <Row label="รายได้จากการขาย" value={`฿${fmt(totalRevenue)}`} />
          <Row label="รายได้อื่น" value={`฿${fmt(totalOtherIncome)}`} />
          {openingApplies && Number(openingRevenue) > 0 && <Row label={`รายได้ยกมา${openingMonth ? " (" + openingMonth + ")" : ""}`} value={`+฿${fmt(Number(openingRevenue))}`} color="#854f0b" />}
          <div style={{ borderTop: "1px solid #e5e7eb", margin: "8px 0" }} />
          <Row label="รวมรายได้" value={`฿${fmt(totalIncome)}`} bold />

          <div style={{ marginTop: 16, marginBottom: 4, fontWeight: 600, fontSize: 13, color: "#6b7280" }}>ต้นทุนขาย:</div>
          <Row label="　สินค้าคงเหลือยกมาต้นงวด" value={`฿${fmt(beginningInventory)}`} />
          <Row label="　บวก ซื้อสินค้า" value={`+฿${fmt(purchasesInRange)}`} />
          {openingApplies && Number(openingCost) > 0 && <Row label={`　ต้นทุนยกมา${openingMonth ? " (" + openingMonth + ")" : ""}`} value={`+฿${fmt(Number(openingCost))}`} color="#854f0b" />}
          <div style={{ borderTop: "1px solid #e5e7eb", margin: "6px 0 6px 16px" }} />
          <Row label="　สินค้าที่มีไว้เพื่อขาย" value={`฿${fmt(goodsAvailableForSale)}`} />
          <Row label="　หัก สินค้าคงเหลือปลายงวด" value={`-฿${fmt(endingInventory)}`} />
          <div style={{ borderTop: "1px solid #e5e7eb", margin: "6px 0 6px 16px" }} />
          <Row label="　ต้นทุนขาย" value={`฿${fmt(totalCost)}`} bold />

          <div style={{ borderTop: "1px solid #e5e7eb", margin: "12px 0" }} />
          <Row label="กำไรขั้นต้น (Gross Profit)" value={`฿${fmt(grossProfit)}`} bold color="#185fa5" />

          <div style={{ marginTop: 12, marginBottom: 4, fontWeight: 600, fontSize: 13, color: "#6b7280" }}>หัก ค่าใช้จ่าย:</div>
          {expenseByCategory.map((c) => (
            <Row key={c.category} label={`　${c.category}`} value={`-฿${fmt(c.amount)}`} />
          ))}
          {expenseByCategory.length === 0 && <Row label="　ไม่มีค่าใช้จ่าย" value="฿0" />}
          <Row label="รวมค่าใช้จ่าย" value={`-฿${fmt(totalExpenses)}`} />
          <div style={{ borderTop: "2px solid #0c443c", margin: "8px 0" }} />
          <Row label="กำไรสุทธิ (Net Profit)" value={`฿${fmt(netProfit)}`} bold color={netProfit >= 0 ? "#0f6e56" : "#993c1d"} />
        </div>

        {/* Dividend Calculation */}
        <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "20px 24px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
            <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700 }}>คำนวณเงินปันผลตามสัดส่วนหุ้น</h3>
            <button style={btnSecondary} onClick={() => setEditingShareholders(!editingShareholders)}>
              <Edit2 size={14} /> {editingShareholders ? "เสร็จสิ้น" : "แก้ไขสัดส่วนหุ้น"}
            </button>
          </div>

          <div style={{ background: "#f9fafb", borderRadius: 8, padding: "12px 16px", marginBottom: 16, fontSize: 14 }}>
            <Row label="กำไรสุทธิที่นำมาคำนวณปันผล" value={`฿${fmt(dividendPool)}`} bold />
            {netProfit < 0 && <p style={{ fontSize: 12, color: "#993c1d", margin: "6px 0 0" }}>* เดือนนี้ขาดทุน ไม่มีเงินปันผล</p>}
          </div>

          {editingShareholders && (
            <div style={{ marginBottom: 12 }}>
              <button style={btnSecondary} onClick={addShareholder}><Plus size={14} /> เพิ่มผู้ถือหุ้น</button>
            </div>
          )}

          <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
            <thead>
              <tr>
                <th style={thStyle}>ชื่อผู้ถือหุ้น/หุ้นส่วน</th>
                <th style={{ ...thStyle, textAlign: "right" }}>สัดส่วน (%)</th>
                <th style={{ ...thStyle, textAlign: "right" }}>เงินปันผลที่ได้รับ</th>
                {editingShareholders && <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>}
              </tr>
            </thead>
            <tbody>
              {shareholders.map((sh, idx) => {
                const amount = totalSharePercent > 0 ? dividendPool * ((Number(sh.percent) || 0) / totalSharePercent) : 0;
                return (
                  <tr key={sh.id}>
                    <td style={tdStyle}>
                      {editingShareholders ? (
                        <input style={inputStyle} value={sh.name} onChange={(e) => updateShareholder(idx, "name", e.target.value)} />
                      ) : sh.name}
                    </td>
                    <td style={{ ...tdStyle, textAlign: "right" }}>
                      {editingShareholders ? (
                        <input type="number" style={{ ...inputStyle, textAlign: "right", width: 90 }} value={sh.percent} onChange={(e) => updateShareholder(idx, "percent", e.target.value)} />
                      ) : `${sh.percent}%`}
                    </td>
                    <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#0f6e56" }}>฿{fmt(amount)}</td>
                    {editingShareholders && (
                      <td style={{ ...tdStyle, textAlign: "right" }}>
                        <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบผู้ถือหุ้น "${sh.name}" ใช่หรือไม่?`, () => removeShareholder(idx))}><Trash2 size={14} /></button>
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr>
                <td style={{ ...tdStyle, fontWeight: 700 }}>รวม</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: totalSharePercent === 100 ? "#0f6e56" : "#993c1d" }}>{totalSharePercent}%</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>฿{fmt(dividendPool)}</td>
                {editingShareholders && <td style={tdStyle}></td>}
              </tr>
            </tfoot>
          </table>
          </div>
          {totalSharePercent !== 100 && (
            <p style={{ fontSize: 12, color: "#993c1d", marginTop: 8 }}>⚠️ สัดส่วนหุ้นรวมต้องเท่ากับ 100% (ปัจจุบัน {totalSharePercent}%)</p>
          )}
        </div>
      </div>
      </>
      )}

      {reportView === "yearly" && (
      <>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 20 }}>
        <select style={{ ...inputStyle, width: 100 }} value={year} onChange={(e) => setYear(Number(e.target.value))}>
          {yearOptions.map((y) => <option key={y} value={y}>ปี {y}</option>)}
        </select>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginBottom: 20 }}>
        <div style={{ background: yearlyNetProfitTotal >= 0 ? "#e6f1fb" : "#faeeda", borderRadius: 12, padding: "16px 18px" }}>
          <div style={{ fontSize: 12, color: yearlyNetProfitTotal >= 0 ? "#185fa5" : "#854f0b", marginBottom: 4 }}>กำไรสุทธิรวมทั้งปี {year}</div>
          <div style={{ fontWeight: 700, fontSize: 22, color: yearlyNetProfitTotal >= 0 ? "#185fa5" : "#854f0b" }}>฿{fmt(yearlyNetProfitTotal)}</div>
        </div>
        <div style={{ background: "#eeedfe", borderRadius: 12, padding: "16px 18px" }}>
          <div style={{ fontSize: 12, color: "#3c3489", marginBottom: 4 }}>จ่ายเงินปันผลไปแล้วในปีนี้</div>
          <div style={{ fontWeight: 700, fontSize: 22, color: "#3c3489" }}>฿{fmt(totalDividendPaidThisYear)}</div>
        </div>
        <div style={{ background: (yearlyNetProfitTotal - totalDividendPaidThisYear) >= 0 ? "#e3f5ea" : "#faeeda", borderRadius: 12, padding: "16px 18px" }}>
          <div style={{ fontSize: 12, color: (yearlyNetProfitTotal - totalDividendPaidThisYear) >= 0 ? "#0f6e56" : "#854f0b", marginBottom: 4 }}>กำไร - เงินปันผล</div>
          <div style={{ fontWeight: 700, fontSize: 22, color: (yearlyNetProfitTotal - totalDividendPaidThisYear) >= 0 ? "#0f6e56" : "#854f0b" }}>฿{fmt(yearlyNetProfitTotal - totalDividendPaidThisYear)}</div>
        </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "20px 24px", marginBottom: 20 }}>
        <h3 style={{ margin: "0 0 16px", fontSize: 15, fontWeight: 700 }}>กำไรสุทธิรายเดือน — ปี {year}</h3>
        <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 560 }}>
          <thead>
            <tr>
              <th style={thStyle}>เดือน</th>
              <th style={{ ...thStyle, textAlign: "right" }}>รวมรายได้</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ต้นทุนขาย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>ค่าใช้จ่าย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>กำไรสุทธิ</th>
            </tr>
          </thead>
          <tbody>
            {yearlyMonths.map((m) => (
              <tr key={m.month}>
                <td style={tdStyle}>{m.label}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(m.totalIncome)}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(m.totalCost)}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>฿{fmt(m.totalExpenses)}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: m.netProfit >= 0 ? "#0f6e56" : "#993c1d" }}>฿{fmt(m.netProfit)}</td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr>
              <td style={{ ...tdStyle, fontWeight: 700 }}>รวมทั้งปี</td>
              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>฿{fmt(yearlyMonths.reduce((s,m)=>s+m.totalIncome,0))}</td>
              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>฿{fmt(yearlyMonths.reduce((s,m)=>s+m.totalCost,0))}</td>
              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700 }}>฿{fmt(yearlyMonths.reduce((s,m)=>s+m.totalExpenses,0))}</td>
              <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: yearlyNetProfitTotal >= 0 ? "#0f6e56" : "#993c1d" }}>฿{fmt(yearlyNetProfitTotal)}</td>
            </tr>
          </tfoot>
        </table>
        </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e5e7eb", padding: "20px 24px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontSize: 15, fontWeight: 700 }}>บันทึกการจ่ายเงินปันผล — ปี {year}</h3>
          <button style={btnPrimary} onClick={openDivPayForm}><Plus size={16} /> บันทึกจ่ายเงินปันผล</button>
        </div>

        <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 400 }}>
          <thead>
            <tr>
              <th style={thStyle}>วันที่จ่าย</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จำนวนเงินที่จ่ายจริง</th>
              <th style={{ ...thStyle, textAlign: "right" }}>จัดการ</th>
            </tr>
          </thead>
          <tbody>
            {dividendPaymentsThisYear.sort((a,b) => (b.date||"").localeCompare(a.date||"")).map((d) => (
              <tr key={d.id}>
                <td style={tdStyle}>{d.date}</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 600, color: "#0f6e56" }}>฿{fmt(d.amount)}</td>
                <td style={{ ...tdStyle, textAlign: "right" }}>
                  <button style={btnDanger} onClick={() => confirmAction(`ต้องการลบรายการจ่ายเงินปันผลวันที่ ${d.date} จำนวน ฿${fmt(d.amount)} ใช่หรือไม่?`, () => removeDivPayment(d.id))}><Trash2 size={14} /> ลบ</button>
                </td>
              </tr>
            ))}
            {dividendPaymentsThisYear.length === 0 && (
              <tr><td colSpan={3} style={{ ...tdStyle, textAlign: "center", color: "#9ca3af" }}>ยังไม่มีการบันทึกจ่ายเงินปันผลในปีนี้</td></tr>
            )}
          </tbody>
          {dividendPaymentsThisYear.length > 0 && (
            <tfoot>
              <tr>
                <td style={{ ...tdStyle, fontWeight: 700 }}>รวม</td>
                <td style={{ ...tdStyle, textAlign: "right", fontWeight: 700, color: "#0f6e56" }}>฿{fmt(totalDividendPaidThisYear)}</td>
                <td style={tdStyle}></td>
              </tr>
            </tfoot>
          )}
        </table>
        </div>

        {/* ยอดยกมาก่อนเริ่มใช้แอพ */}
        <div style={{ background: "#fffbeb", borderRadius: 12, border: "1px solid #fde68a", padding: "16px 20px", marginTop: 16 }}>
          <div style={{ fontWeight: 600, fontSize: 14, color: "#854f0b", marginBottom: 12 }}>ยอดยกมาก่อนเริ่มใช้แอพ (ถ้ามี)</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 16px" }}>
            <div>
              <label style={{ display: "block", fontSize: 13, color: "#374151", marginBottom: 4 }}>เดือนที่มีผล</label>
              <input type="month" style={inputStyle} value={openingMonth} onChange={(e) => setOpeningMonth(e.target.value)} />
            </div>
            <div>
              <label style={{ display: "block", fontSize: 13, color: "#374151", marginBottom: 4 }}>รายได้ยกมา (บาท)</label>
              <input type="number" style={{ ...inputStyle, textAlign: "right" }} value={openingRevenue} onChange={(e) => setOpeningRevenue(e.target.value)} placeholder="0" />
            </div>
            <div>
              <label style={{ display: "block", fontSize: 13, color: "#374151", marginBottom: 4 }}>ต้นทุนยกมา (บาท)</label>
              <input type="number" style={{ ...inputStyle, textAlign: "right" }} value={openingCost} onChange={(e) => setOpeningCost(e.target.value)} placeholder="0" />
            </div>
          </div>
          <p style={{ fontSize: 12, color: "#854f0b", margin: "8px 0 0" }}>* ยอดยกมาจะรวมเข้างบเฉพาะเดือนที่กำหนด — ถ้าไม่กำหนดเดือนจะรวมทุกเดือน — บันทึกถาวรอัตโนมัติ ใช้ได้ทุกเครื่อง</p>
        </div>

        {divPayForm && (
          <Modal title="บันทึกจ่ายเงินปันผล" onClose={() => setDivPayForm(null)}>
            <Field label="วันที่จ่าย">
              <input type="date" style={inputStyle} value={divPayForm.date} onChange={(e) => setDivPayForm({ ...divPayForm, date: e.target.value })} />
            </Field>
            <Field label="จำนวนเงินที่จ่ายจริง (รวมทุกคน)">
              <input type="number" style={inputStyle} value={divPayForm.amount} onChange={(e) => setDivPayForm({ ...divPayForm, amount: e.target.value })} placeholder="0" />
            </Field>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 16 }}>
              <button style={btnSecondary} onClick={() => setDivPayForm(null)}>ยกเลิก</button>
              <button style={btnPrimary} onClick={saveDivPayment}><Save size={16} /> บันทึก</button>
            </div>
          </Modal>
        )}
      </div>
      </>
      )}
    </div>
  );
}
