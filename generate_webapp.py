import os
import base64
import json
import pandas as pd

def encode_img_to_b64(path):
    if os.path.exists(path):
        with open(path, "rb") as image_file:
            return "data:image/jpeg;base64," + base64.b64encode(image_file.read()).decode('utf-8')
    return ""

def encode_png_to_b64(path):
    if os.path.exists(path):
        with open(path, "rb") as image_file:
            return "data:image/png;base64," + base64.b64encode(image_file.read()).decode('utf-8')
    return ""

excel_file = r"G:\Exium\2026\4Q'26\Sweater\FF list.xlsx"
df = pd.read_excel(excel_file)
territories = df.to_dict(orient='records')

region_map = {}
for t in territories:
    r_code = str(t['SAP Region Code']).strip()
    if r_code not in region_map:
        region_map[r_code] = {
            "sap_region_code": r_code,
            "region_name": str(t['Region']).strip(),
            "zone": str(t['Zone']).strip(),
            "regional_head": str(t['Regional Head']).strip(),
            "territories": []
        }
    region_map[r_code]["territories"].append({
        "sap_territory_code": str(t['SAP Territory Code']).strip(),
        "territory_name": str(t['Territory']).strip()
    })

zones = sorted(list(set(t['Zone'] for t in territories)))
default_cloud_url = "https://script.google.com/macros/s/AKfycbzEnDTtNiXEAyB5qHqrxLj1RbNytgOJAB_lKjw_VVVd1C8CiaeYU6iTROiJabkyX_-b/exec"

print("Encoding images...")
b64_logo_sq = encode_png_to_b64(r"G:\Exium\2026\4Q'26\Sweater\Square Logo.png")
b64_logo_banner = encode_png_to_b64(r"G:\Exium\2026\4Q'26\Sweater\Exium MUPS Logo.png")
b64_01 = encode_img_to_b64(r"G:\Exium\2026\4Q'26\Sweater\Image\01 (Men).jpeg")
b64_02 = encode_img_to_b64(r"G:\Exium\2026\4Q'26\Sweater\Image\02 (Men).jpeg")
b64_03 = encode_img_to_b64(r"G:\Exium\2026\4Q'26\Sweater\Image\03 (Men).jpeg")
b64_04 = encode_img_to_b64(r"G:\Exium\2026\4Q'26\Sweater\Image\04 (Female).jpeg")
b64_05 = encode_img_to_b64(r"G:\Exium\2026\4Q'26\Sweater\Image\05 (Female).jpeg")

html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exium MUPS - 4Q'26 Sweater Campaign Portal</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome CDN -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS (XLSX) CDN -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        .custom-scrollbar::-webkit-scrollbar { height: 6px; width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 9999px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 9999px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        .animate-fade-in { animation: fadeIn 0.25s cubic-bezier(0.16, 1, 0.3, 1) forwards; }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen font-sans antialiased flex flex-col selection:bg-orange-500 selection:text-white">

    <!-- TOP GLOBAL HEADER -->
    <header class="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-sm backdrop-blur-md bg-white/90">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 py-2.5 sm:py-3.5 flex items-center justify-between gap-3">
            <div class="flex items-center gap-2.5 sm:gap-4">
                <img src="###B64_LOGO_SQ###" onerror="this.src='Square Logo.png'" alt="Logo" class="w-8 h-8 sm:w-10 sm:h-10 object-contain rounded-lg shadow-sm border border-slate-100 p-0.5">
                <div class="leading-tight">
                    <div class="flex items-center gap-1.5 sm:gap-2">
                        <h1 class="text-sm sm:text-lg font-black text-slate-900 tracking-tight">EXIUM <span class="text-orange-600">MUPS</span></h1>
                        <span class="px-2 py-0.5 bg-orange-100 text-orange-800 text-[10px] sm:text-xs font-black rounded-full border border-orange-200 uppercase tracking-wide">4Q'26</span>
                    </div>
                    <p class="text-[10px] sm:text-xs text-slate-500 font-medium hidden sm:block">Doctor Sweater Campaign Portal • Radiant Pharmaceuticals</p>
                </div>
            </div>

            <!-- Header Action Buttons -->
            <div class="flex items-center gap-2">
                <!-- Live Dhaka Time (Desktop) -->
                <div class="hidden lg:flex items-center gap-1.5 px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-600">
                    <i class="fa-regular fa-clock text-orange-500"></i>
                    <span id="live-dhaka-clock">Dhaka Time</span>
                </div>

                <!-- Catalogue & Sizes Button -->
                <button onclick="openCatalogModal()" class="px-3 sm:px-4 py-1.5 sm:py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs sm:text-sm font-bold rounded-xl flex items-center gap-1.5 sm:gap-2 shadow-sm transition active:scale-95">
                    <i class="fa-solid fa-vest text-orange-400"></i>
                    <span class="hidden sm:inline">Catalogue & Sizes</span>
                    <span class="sm:hidden">Catalogue</span>
                </button>

                <!-- Admin Dashboard Button -->
                <button onclick="openAdminModal()" class="px-3 sm:px-4 py-1.5 sm:py-2 bg-orange-600 hover:bg-orange-500 text-white text-xs sm:text-sm font-bold rounded-xl flex items-center gap-1.5 sm:gap-2 shadow-sm transition active:scale-95">
                    <i class="fa-solid fa-chart-pie"></i>
                    <span class="hidden sm:inline">Admin Portal</span>
                    <span class="sm:hidden">Admin</span>
                </button>
            </div>
        </div>
    </header>

    <!-- MAIN APP WRAPPER -->
    <main class="flex-1 max-w-7xl w-full mx-auto p-3 sm:p-6 space-y-4 sm:space-y-6">

        <!-- ============================================== -->
        <!-- VIEW 1: REGION LOGIN / SELECTOR VIEW          -->
        <!-- ============================================== -->
        <section id="selection-view" class="bg-white border border-slate-200 rounded-3xl p-5 sm:p-8 shadow-sm transition">
            <div class="max-w-2xl mx-auto space-y-6">
                <div class="text-center space-y-2">
                    <div class="inline-flex p-3 rounded-2xl bg-orange-50 text-orange-600 text-2xl font-black mb-1 border border-orange-100 shadow-sm">
                        <i class="fa-solid fa-map-location-dot"></i>
                    </div>
                    <h2 class="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">Regional Manager Login</h2>
                    <p class="text-xs sm:text-sm text-slate-500">Select your Zone and Region to access your territory doctor allocations</p>
                </div>

                <div class="space-y-4 bg-slate-50 p-4 sm:p-6 rounded-2xl border border-slate-200">
                    <!-- Step 1: Select Zone -->
                    <div class="space-y-1.5">
                        <label class="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                            <span class="w-4 h-4 rounded-full bg-slate-900 text-white text-[10px] flex items-center justify-center font-bold">1</span>
                            <span>Select Zone</span>
                        </label>
                        <select id="select-zone" onchange="onZoneChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition shadow-sm">
                            <option value="">-- Choose Your Zone (35 Zones) --</option>
                        </select>
                    </div>

                    <!-- Step 2: Select Region -->
                    <div class="space-y-1.5">
                        <label class="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                            <span class="w-4 h-4 rounded-full bg-slate-900 text-white text-[10px] flex items-center justify-center font-bold">2</span>
                            <span>Select Region</span>
                        </label>
                        <select id="select-region" onchange="onRegionChanged()" disabled class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition disabled:opacity-50 disabled:bg-slate-100 shadow-sm">
                            <option value="">-- First Select Zone Above --</option>
                        </select>
                    </div>

                    <!-- Selected Region Details Card -->
                    <div id="rh-info-card" class="hidden bg-white border border-orange-200 rounded-2xl p-4 space-y-2 animate-fade-in shadow-sm">
                        <div class="flex items-center justify-between border-b border-slate-100 pb-2">
                            <span class="text-[11px] font-bold uppercase text-slate-400">Regional Information</span>
                            <span id="card-sap-badge" class="px-2 py-0.5 rounded-full text-[10px] font-black bg-orange-50 text-orange-700 border border-orange-200">SAP: 00000</span>
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                            <div>
                                <span class="text-[10px] text-slate-400 block font-medium">Region Name</span>
                                <strong id="card-region-name" class="font-bold text-slate-900 text-sm">--</strong>
                            </div>
                            <div>
                                <span class="text-[10px] text-slate-400 block font-medium">Regional Head</span>
                                <strong id="card-rh-name" class="font-bold text-slate-900 text-sm">--</strong>
                            </div>
                        </div>
                    </div>

                    <!-- Step 3: Password Section -->
                    <div id="password-section" class="hidden space-y-1.5 animate-fade-in">
                        <label class="text-xs font-bold text-slate-700 flex items-center justify-between">
                            <span class="flex items-center gap-1.5">
                                <span class="w-4 h-4 rounded-full bg-slate-900 text-white text-[10px] flex items-center justify-center font-bold">3</span>
                                <span>Security Password</span>
                            </span>
                            <span class="text-[10px] text-slate-400 font-normal">Use SAP Region Code or Default Pass</span>
                        </label>
                        <div class="relative">
                            <input type="password" id="region-password" placeholder="Enter password to unlock..." onkeyup="handlePasswordKey(event)" class="w-full bg-white border border-slate-300 rounded-xl pl-3.5 pr-10 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition shadow-sm">
                            <button type="button" onclick="togglePasswordVisibility()" class="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-sm"><i class="fa-regular fa-eye" id="pass-toggle-icon"></i></button>
                        </div>
                    </div>

                    <!-- Step 4: Login Action Button -->
                    <div id="unlock-btn-container" class="hidden pt-2 animate-fade-in">
                        <button onclick="unlockRegion()" class="w-full py-3 bg-gradient-to-r from-orange-600 to-amber-600 hover:from-orange-500 hover:to-amber-500 text-white text-sm font-black rounded-xl flex items-center justify-center gap-2 shadow-md transition transform active:scale-98">
                            <i class="fa-solid fa-lock-open"></i>
                            <span>Access Workspace</span>
                        </button>
                    </div>
                </div>
            </div>
        </section>

        <!-- ============================================== -->
        <!-- VIEW 2: REGIONAL MANAGER WORKSPACE VIEW       -->
        <!-- ============================================== -->
        <section id="workspace-view" class="hidden space-y-4 sm:space-y-6 animate-fade-in">

            <!-- Regional Banner Card -->
            <div class="bg-white border border-slate-200 rounded-3xl p-4 sm:p-6 shadow-sm flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                <div class="space-y-1">
                    <div class="flex items-center gap-2 flex-wrap">
                        <span id="banner-region" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200">SAP: 00000</span>
                        <span id="banner-zone" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-orange-50 text-orange-700 border border-orange-200">Zone</span>
                    </div>
                    <h2 id="banner-rh" class="text-base sm:text-xl font-black text-slate-900 tracking-tight">Region: Region Name (Regional Head Name)</h2>
                    <p class="text-xs text-slate-500 font-medium">Please fill allocations for all territories in your region. Data auto-syncs live.</p>
                </div>

                <!-- Action Controls -->
                <div class="flex items-center gap-2 sm:gap-3 flex-wrap">
                    <!-- Progress Badge -->
                    <div class="bg-slate-50 border border-slate-200 rounded-2xl px-3.5 py-2 flex items-center gap-2.5">
                        <div class="text-right">
                            <div class="text-[10px] text-slate-400 font-bold uppercase">Region Progress</div>
                            <div id="region-progress-text" class="text-xs font-black text-slate-900">0 / 0 Complete</div>
                        </div>
                        <div class="w-8 h-8 rounded-xl bg-orange-100 text-orange-700 flex items-center justify-center font-black text-xs" id="region-progress-pct">0%</div>
                    </div>

                    <!-- Export Single Region Excel -->
                    <button onclick="exportCurrentRegionExcel()" class="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black rounded-xl flex items-center gap-1.5 shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-file-excel"></i>
                        <span class="hidden sm:inline">Export Excel</span>
                    </button>

                    <!-- Exit / Switch Region -->
                    <button onclick="exitRegionWorkspace()" class="px-3.5 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-bold rounded-xl flex items-center gap-1.5 transition active:scale-95">
                        <i class="fa-solid fa-right-from-bracket"></i>
                        <span>Exit</span>
                    </button>

                    <!-- Save Territory Button -->
                    <button onclick="saveCurrentTerritoryClick()" class="px-5 py-2 bg-orange-600 hover:bg-orange-500 text-white text-xs font-black rounded-xl flex items-center gap-2 shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-floppy-disk"></i>
                        <span>Save</span>
                    </button>
                </div>
            </div>

            <!-- Territory Navigation Tabs (Desktop Horizontal Scroll + Mobile Dropdown) -->
            <div class="bg-white border border-slate-200 rounded-2xl p-2 sm:p-3 shadow-sm space-y-2">
                <!-- Mobile Territory Dropdown Selector -->
                <div class="block sm:hidden">
                    <label class="text-[11px] font-bold text-slate-500 uppercase tracking-wide mb-1 block">Select Territory</label>
                    <select id="mobile-territory-select" onchange="selectTerritoryTab(parseInt(this.value))" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800">
                    </select>
                </div>

                <!-- Desktop Horizontal Tabs -->
                <div id="desktop-territory-tabs" class="hidden sm:flex items-center gap-2 overflow-x-auto custom-scrollbar pb-1">
                </div>
            </div>

            <!-- ============================================== -->
            <!-- ACTIVE TERRITORY FORM SECTION                 -->
            <!-- ============================================== -->
            <div class="grid grid-cols-1 gap-6">

                <!-- ACTIVE TERRITORY HEADER -->
                <div class="bg-slate-900 text-white rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm">
                    <div class="space-y-0.5">
                        <div class="flex items-center gap-2">
                            <span id="current-territory-code" class="text-[10px] font-mono font-bold bg-slate-800 text-orange-400 px-2 py-0.5 rounded border border-slate-700">SAP Code: 00000</span>
                            <span id="current-territory-status" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-slate-800 text-slate-300">Not Started</span>
                        </div>
                        <h3 id="current-territory-title" class="text-lg sm:text-xl font-black text-white tracking-tight">Territory Name</h3>
                    </div>

                    <div class="flex items-center gap-2 text-xs">
                        <span class="text-slate-400 font-medium">Auto-sync:</span>
                        <span id="auto-sync-status" class="flex items-center gap-1.5 font-bold text-emerald-400">
                            <span class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                            <span>Live Cloud Sync Ready</span>
                        </span>
                    </div>
                </div>

                <!-- CAMPAIGN 1: GYNE CORE DOCTOR DEVELOPMENT (FAMILY PACKAGE) -->
                <div class="bg-white border-2 border-teal-500/30 rounded-3xl p-4 sm:p-6 shadow-sm space-y-5">
                    <!-- Campaign 1 Title Header -->
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-teal-100 gap-2 bg-teal-50/50 -m-4 sm:-m-6 p-4 sm:p-6 rounded-t-3xl border-b border-teal-200">
                        <div class="flex items-center gap-3">
                            <div class="w-10 h-10 rounded-2xl bg-teal-600 text-white flex items-center justify-center font-black text-lg shadow-sm">1</div>
                            <div>
                                <h4 class="font-black text-slate-900 text-sm sm:text-base tracking-tight">CAMPAIGN 1: GYNE CORE DOCTOR DEVELOPMENT</h4>
                                <p class="text-[11px] text-teal-800 font-medium">Family Package • 1 Doctor gets 4 Sweaters (Men's / Women's as requested)</p>
                            </div>
                        </div>
                        <span class="self-start sm:self-auto px-3 py-1 bg-teal-100 text-teal-900 border border-teal-300 rounded-full text-[10px] font-black uppercase tracking-wider">Family Bundle (4 Sweaters)</span>
                    </div>

                    <!-- Doctor Information (Gyne Core Doctor) -->
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 bg-slate-50 p-4 rounded-2xl border border-slate-200">
                        <div class="space-y-1">
                            <label class="text-xs font-bold text-slate-700">Doctor Name <span class="text-rose-500">*</span></label>
                            <input type="text" id="c1_doc_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Enter Full Doctor Name..." class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 focus:outline-none transition">
                        </div>
                        <div class="space-y-1">
                            <div class="flex items-center justify-between">
                                <label class="text-xs font-bold text-slate-700">Doctor RPL ID (6 Digits) <span class="text-rose-500">*</span></label>
                                <span id="c1_doc_rpl_badge" class="text-[10px] font-bold text-slate-400">6 digits</span>
                            </div>
                            <input type="text" inputmode="numeric" maxlength="6" id="c1_doc_rpl" oninput="onRplInput(this, 'c1_doc_rpl_badge')" onchange="onRplInput(this, 'c1_doc_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c1_doc_rpl_badge'), 50)" onblur="onRplInput(this, 'c1_doc_rpl_badge')" placeholder="e.g. 104523" class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 font-mono font-bold placeholder-slate-400 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 focus:outline-none transition tracking-wider">
                            <div id="c1_doc_rpl_dup_msg" class="hidden"></div>
                        </div>
                    </div>

                    <!-- 4 Sweaters Allocation for Doctor & Family -->
                    <div class="space-y-3">
                        <div class="flex items-center justify-between">
                            <label class="text-xs font-black text-slate-900 uppercase tracking-wide">4 Sweaters Selection (Family Allocation)</label>
                            <span class="text-[10px] text-slate-500 font-medium">Click thumbnail to zoom design</span>
                        </div>

                        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3.5">
                            <!-- Sweater 1 -->
                            <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2.5 relative">
                                <div class="flex items-center justify-between">
                                    <span class="text-[11px] font-bold text-teal-800">Sweater 1</span>
                                    <div id="c1_m1_img_thumb" class="w-7 h-9 bg-white border border-slate-200 rounded overflow-hidden cursor-pointer shadow-2xs" onclick="openSlotLightbox('c1_m1_sweater')">
                                        <img src="" class="w-full h-full object-cover hidden" alt="Thumb">
                                    </div>
                                </div>
                                <select id="c1_m1_sweater" onchange="onDataChanged(); updateSlotThumbnail('c1_m1_sweater', 'c1_m1_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                                <select id="c1_m1_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>

                            <!-- Sweater 2 -->
                            <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2.5 relative">
                                <div class="flex items-center justify-between">
                                    <span class="text-[11px] font-bold text-teal-800">Sweater 2</span>
                                    <div id="c1_m2_img_thumb" class="w-7 h-9 bg-white border border-slate-200 rounded overflow-hidden cursor-pointer shadow-2xs" onclick="openSlotLightbox('c1_m2_sweater')">
                                        <img src="" class="w-full h-full object-cover hidden" alt="Thumb">
                                    </div>
                                </div>
                                <select id="c1_m2_sweater" onchange="onDataChanged(); updateSlotThumbnail('c1_m2_sweater', 'c1_m2_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                                <select id="c1_m2_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>

                            <!-- Sweater 3 -->
                            <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2.5 relative">
                                <div class="flex items-center justify-between">
                                    <span class="text-[11px] font-bold text-teal-800">Sweater 3</span>
                                    <div id="c1_m3_img_thumb" class="w-7 h-9 bg-white border border-slate-200 rounded overflow-hidden cursor-pointer shadow-2xs" onclick="openSlotLightbox('c1_m3_sweater')">
                                        <img src="" class="w-full h-full object-cover hidden" alt="Thumb">
                                    </div>
                                </div>
                                <select id="c1_m3_sweater" onchange="onDataChanged(); updateSlotThumbnail('c1_m3_sweater', 'c1_m3_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                                <select id="c1_m3_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>

                            <!-- Sweater 4 -->
                            <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2.5 relative">
                                <div class="flex items-center justify-between">
                                    <span class="text-[11px] font-bold text-teal-800">Sweater 4</span>
                                    <div id="c1_m4_img_thumb" class="w-7 h-9 bg-white border border-slate-200 rounded overflow-hidden cursor-pointer shadow-2xs" onclick="openSlotLightbox('c1_m4_sweater')">
                                        <img src="" class="w-full h-full object-cover hidden" alt="Thumb">
                                    </div>
                                </div>
                                <select id="c1_m4_sweater" onchange="onDataChanged(); updateSlotThumbnail('c1_m4_sweater', 'c1_m4_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                                <select id="c1_m4_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- CAMPAIGN 2: CORE DOCTOR MAXIMIZATION (4 DOCTORS PER TERRITORY) -->
                <div class="bg-white border-2 border-purple-500/30 rounded-3xl p-4 sm:p-6 shadow-sm space-y-5">
                    <!-- Campaign 2 Title Header -->
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-3 border-b border-purple-100 gap-2 bg-purple-50/50 -m-4 sm:-m-6 p-4 sm:p-6 rounded-t-3xl border-b border-purple-200">
                        <div class="flex items-center gap-3">
                            <div class="w-10 h-10 rounded-2xl bg-purple-700 text-white flex items-center justify-center font-black text-lg shadow-sm">2</div>
                            <div>
                                <h4 class="font-black text-slate-900 text-sm sm:text-base tracking-tight">CAMPAIGN 2: CORE DOCTOR MAXIMIZATION</h4>
                                <p class="text-[11px] text-purple-800 font-medium">Individual Allocation • 4 Core Doctors (1 Sweater each / territory)</p>
                            </div>
                        </div>
                        <span class="self-start sm:self-auto px-3 py-1 bg-purple-100 text-purple-900 border border-purple-300 rounded-full text-[10px] font-black uppercase tracking-wider">4 Core Doctors</span>
                    </div>

                    <!-- 4 Doctor Cards Grid -->
                    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">

                        <!-- Doctor 1 -->
                        <div class="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-3">
                            <div class="flex items-center justify-between">
                                <span class="w-6 h-6 rounded-lg bg-purple-700 text-white text-xs font-black flex items-center justify-center">1</span>
                                <span class="text-xs font-black text-purple-900">Doctor 1</span>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Doctor Name <span class="text-rose-500">*</span></label>
                                <input type="text" id="c2_d1_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Doctor 1 Name..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-semibold text-slate-800">
                            </div>
                            <div class="space-y-1">
                                <div class="flex items-center justify-between">
                                    <label class="text-[10px] font-bold text-slate-500 uppercase">RPL ID (6 Digits) <span class="text-rose-500">*</span></label>
                                    <span id="c2_d1_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span>
                                </div>
                                <input type="text" inputmode="numeric" maxlength="6" id="c2_d1_rpl" oninput="onRplInput(this, 'c2_d1_rpl_badge')" onchange="onRplInput(this, 'c2_d1_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d1_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d1_rpl_badge')" placeholder="6-digit RPL..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-mono font-bold text-slate-800">
                                <div id="c2_d1_rpl_dup_msg" class="hidden"></div>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Sweater Design <span class="text-rose-500">*</span></label>
                                <select id="c2_d1_sweater" onchange="onDataChanged(); updateSlotThumbnail('c2_d1_sweater', 'c2_d1_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Size <span class="text-rose-500">*</span></label>
                                <select id="c2_d1_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>
                        </div>

                        <!-- Doctor 2 -->
                        <div class="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-3">
                            <div class="flex items-center justify-between">
                                <span class="w-6 h-6 rounded-lg bg-purple-700 text-white text-xs font-black flex items-center justify-center">2</span>
                                <span class="text-xs font-black text-purple-900">Doctor 2</span>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Doctor Name <span class="text-rose-500">*</span></label>
                                <input type="text" id="c2_d2_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Doctor 2 Name..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-semibold text-slate-800">
                            </div>
                            <div class="space-y-1">
                                <div class="flex items-center justify-between">
                                    <label class="text-[10px] font-bold text-slate-500 uppercase">RPL ID (6 Digits) <span class="text-rose-500">*</span></label>
                                    <span id="c2_d2_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span>
                                </div>
                                <input type="text" inputmode="numeric" maxlength="6" id="c2_d2_rpl" oninput="onRplInput(this, 'c2_d2_rpl_badge')" onchange="onRplInput(this, 'c2_d2_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d2_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d2_rpl_badge')" placeholder="6-digit RPL..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-mono font-bold text-slate-800">
                                <div id="c2_d2_rpl_dup_msg" class="hidden"></div>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Sweater Design <span class="text-rose-500">*</span></label>
                                <select id="c2_d2_sweater" onchange="onDataChanged(); updateSlotThumbnail('c2_d2_sweater', 'c2_d2_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Size <span class="text-rose-500">*</span></label>
                                <select id="c2_d2_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>
                        </div>

                        <!-- Doctor 3 -->
                        <div class="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-3">
                            <div class="flex items-center justify-between">
                                <span class="w-6 h-6 rounded-lg bg-purple-700 text-white text-xs font-black flex items-center justify-center">3</span>
                                <span class="text-xs font-black text-purple-900">Doctor 3</span>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Doctor Name <span class="text-rose-500">*</span></label>
                                <input type="text" id="c2_d3_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Doctor 3 Name..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-semibold text-slate-800">
                            </div>
                            <div class="space-y-1">
                                <div class="flex items-center justify-between">
                                    <label class="text-[10px] font-bold text-slate-500 uppercase">RPL ID (6 Digits) <span class="text-rose-500">*</span></label>
                                    <span id="c2_d3_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span>
                                </div>
                                <input type="text" inputmode="numeric" maxlength="6" id="c2_d3_rpl" oninput="onRplInput(this, 'c2_d3_rpl_badge')" onchange="onRplInput(this, 'c2_d3_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d3_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d3_rpl_badge')" placeholder="6-digit RPL..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-mono font-bold text-slate-800">
                                <div id="c2_d3_rpl_dup_msg" class="hidden"></div>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Sweater Design <span class="text-rose-500">*</span></label>
                                <select id="c2_d3_sweater" onchange="onDataChanged(); updateSlotThumbnail('c2_d3_sweater', 'c2_d3_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Size <span class="text-rose-500">*</span></label>
                                <select id="c2_d3_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>
                        </div>

                        <!-- Doctor 4 -->
                        <div class="bg-slate-50 border border-slate-200 rounded-2xl p-4 space-y-3">
                            <div class="flex items-center justify-between">
                                <span class="w-6 h-6 rounded-lg bg-purple-700 text-white text-xs font-black flex items-center justify-center">4</span>
                                <span class="text-xs font-black text-purple-900">Doctor 4</span>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Doctor Name <span class="text-rose-500">*</span></label>
                                <input type="text" id="c2_d4_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Doctor 4 Name..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-semibold text-slate-800">
                            </div>
                            <div class="space-y-1">
                                <div class="flex items-center justify-between">
                                    <label class="text-[10px] font-bold text-slate-500 uppercase">RPL ID (6 Digits) <span class="text-rose-500">*</span></label>
                                    <span id="c2_d4_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span>
                                </div>
                                <input type="text" inputmode="numeric" maxlength="6" id="c2_d4_rpl" oninput="onRplInput(this, 'c2_d4_rpl_badge')" onchange="onRplInput(this, 'c2_d4_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d4_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d4_rpl_badge')" placeholder="6-digit RPL..." class="w-full bg-white border border-slate-300 rounded-xl px-3 py-1.5 text-xs font-mono font-bold text-slate-800">
                                <div id="c2_d4_rpl_dup_msg" class="hidden"></div>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Sweater Design <span class="text-rose-500">*</span></label>
                                <select id="c2_d4_sweater" onchange="onDataChanged(); updateSlotThumbnail('c2_d4_sweater', 'c2_d4_img_thumb')" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-semibold text-slate-800">
                                    <option value="">-- Select Design --</option>
                                </select>
                            </div>
                            <div class="space-y-1">
                                <label class="text-[10px] font-bold text-slate-500 uppercase">Size <span class="text-rose-500">*</span></label>
                                <select id="c2_d4_size" onchange="onDataChanged()" class="w-full bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-800">
                                    <option value="">-- Select Size --</option>
                                </select>
                            </div>
                        </div>

                    </div>
                </div>

                <!-- BOTTOM SAVE BAR -->
                <div class="bg-white border border-slate-200 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm">
                    <div class="flex items-center gap-2 text-xs text-slate-500">
                        <i class="fa-solid fa-circle-info text-orange-500 text-sm"></i>
                        <span>Changes are automatically saved locally and synced to Google Sheet when valid.</span>
                    </div>
                    <button onclick="saveCurrentTerritoryClick()" class="px-6 py-2.5 bg-orange-600 hover:bg-orange-500 text-white font-black text-sm rounded-xl flex items-center justify-center gap-2 shadow-md transition active:scale-98">
                        <i class="fa-solid fa-floppy-disk"></i>
                        <span>Save & Sync Territory</span>
                    </button>
                </div>

            </div>
        </section>

    </main>

    <!-- ============================================== -->
    <!-- CATALOGUE & SIZES MODAL                       -->
    <!-- ============================================== -->
    <div id="catalog-modal" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-3 sm:p-6" onclick="closeCatalogModal()">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden" onclick="event.stopPropagation()">
            <div class="px-5 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
                <div class="flex items-center gap-2.5">
                    <div class="w-8 h-8 rounded-xl bg-orange-500 text-white flex items-center justify-center font-bold text-sm shadow-sm"><i class="fa-solid fa-vest"></i></div>
                    <div>
                        <h3 class="font-bold text-xs sm:text-sm text-slate-900">Sweater Designs Catalogue & Size Specification</h3>
                        <p class="text-[10px] sm:text-xs text-slate-500">Lubnan Trade Consortium Ltd. (Richman / Lubnan)</p>
                    </div>
                </div>
                <button onclick="closeCatalogModal()" class="w-8 h-8 rounded-full bg-white border border-slate-300 text-slate-600 hover:bg-slate-100 flex items-center justify-center transition"><i class="fa-solid fa-xmark text-sm"></i></button>
            </div>
            
            <div class="p-4 sm:p-6 overflow-y-auto custom-scrollbar space-y-6">
                <!-- 5 Sweater Image Cards -->
                <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3 sm:gap-4">
                    <div onclick="openImageLightbox('01')" class="bg-slate-50 border border-slate-200 hover:border-orange-400 rounded-2xl p-2.5 space-y-2 cursor-pointer transition group shadow-sm">
                        <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden border border-slate-200 relative"><img src="###B64_01###" onerror="this.src='Image/01 (Men).jpeg'" alt="01" class="w-full h-full object-cover"><span class="absolute top-1.5 left-1.5 bg-slate-900/90 text-white text-[10px] font-black px-2 py-0.5 rounded-md">01</span></div>
                        <div><span class="text-[9px] font-bold uppercase text-orange-700 bg-orange-50 px-1.5 py-0.5 rounded">Men's</span><h4 class="text-xs font-black text-slate-900 mt-1 leading-tight">V-Neck (Grey)</h4><p class="text-[10px] text-slate-500">Sizes: S - XXL</p></div>
                    </div>
                    <div onclick="openImageLightbox('02')" class="bg-slate-50 border border-slate-200 hover:border-orange-400 rounded-2xl p-2.5 space-y-2 cursor-pointer transition group shadow-sm">
                        <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden border border-slate-200 relative"><img src="###B64_02###" onerror="this.src='Image/02 (Men).jpeg'" alt="02" class="w-full h-full object-cover"><span class="absolute top-1.5 left-1.5 bg-slate-900/90 text-white text-[10px] font-black px-2 py-0.5 rounded-md">02</span></div>
                        <div><span class="text-[9px] font-bold uppercase text-orange-700 bg-orange-50 px-1.5 py-0.5 rounded">Men's</span><h4 class="text-xs font-black text-slate-900 mt-1 leading-tight">V-Neck (Navy)</h4><p class="text-[10px] text-slate-500">Sizes: S - XXL</p></div>
                    </div>
                    <div onclick="openImageLightbox('03')" class="bg-slate-50 border border-slate-200 hover:border-orange-400 rounded-2xl p-2.5 space-y-2 cursor-pointer transition group shadow-sm">
                        <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden border border-slate-200 relative"><img src="###B64_03###" onerror="this.src='Image/03 (Men).jpeg'" alt="03" class="w-full h-full object-cover"><span class="absolute top-1.5 left-1.5 bg-slate-900/90 text-white text-[10px] font-black px-2 py-0.5 rounded-md">03</span></div>
                        <div><span class="text-[9px] font-bold uppercase text-orange-700 bg-orange-50 px-1.5 py-0.5 rounded">Men's</span><h4 class="text-xs font-black text-slate-900 mt-1 leading-tight">V-Neck (Cream)</h4><p class="text-[10px] text-slate-500">Sizes: S - XXL</p></div>
                    </div>
                    <div onclick="openImageLightbox('04')" class="bg-slate-50 border border-slate-200 hover:border-orange-400 rounded-2xl p-2.5 space-y-2 cursor-pointer transition group shadow-sm">
                        <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden border border-slate-200 relative"><img src="###B64_04###" onerror="this.src='Image/04 (Female).jpeg'" alt="04" class="w-full h-full object-cover"><span class="absolute top-1.5 left-1.5 bg-slate-900/90 text-white text-[10px] font-black px-2 py-0.5 rounded-md">04</span></div>
                        <div><span class="text-[9px] font-bold uppercase text-purple-700 bg-purple-50 px-1.5 py-0.5 rounded">Women's</span><h4 class="text-xs font-black text-slate-900 mt-1 leading-tight">Short Cardigan</h4><p class="text-[10px] text-slate-500">Sizes: XS - XL</p></div>
                    </div>
                    <div onclick="openImageLightbox('05')" class="bg-slate-50 border border-slate-200 hover:border-orange-400 rounded-2xl p-2.5 space-y-2 cursor-pointer transition group shadow-sm col-span-2 sm:col-span-1">
                        <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden border border-slate-200 relative"><img src="###B64_05###" onerror="this.src='Image/05 (Female).jpeg'" alt="05" class="w-full h-full object-cover"><span class="absolute top-1.5 left-1.5 bg-slate-900/90 text-white text-[10px] font-black px-2 py-0.5 rounded-md">05</span></div>
                        <div><span class="text-[9px] font-bold uppercase text-purple-700 bg-purple-50 px-1.5 py-0.5 rounded">Women's</span><h4 class="text-xs font-black text-slate-900 mt-1 leading-tight">Semi Long Cardigan</h4><p class="text-[10px] text-slate-500">Sizes: S - XXL</p></div>
                    </div>
                </div>

                <!-- Size & Measurement Charts -->
                <div class="space-y-4 pt-2 border-t border-slate-200">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h4 class="text-sm sm:text-base font-black text-slate-900 flex items-center gap-2">
                                <i class="fa-solid fa-ruler-combined text-orange-500"></i>
                                <span>Sweater Size Measurement Chart (Inches)</span>
                            </h4>
                            <p class="text-[11px] text-slate-500">Standard apparel measurement specifications by Lubnan Trade Consortium Ltd. (Richman / Lubnan)</p>
                        </div>
                        <span class="text-[10px] font-bold bg-orange-100 text-orange-800 border border-orange-200 px-2.5 py-1 rounded-full self-start sm:self-auto">
                            All Measurements in Inches (")
                        </span>
                    </div>

                    <!-- 1. Men's Sleeveless V-Neck Sweaters -->
                    <div class="bg-slate-50 rounded-2xl p-4 border border-slate-200 space-y-2.5">
                        <div class="flex items-center justify-between">
                            <h5 class="text-xs font-black text-slate-900 flex items-center gap-1.5">
                                <span class="w-5 h-5 rounded-md bg-orange-500 text-white text-[10px] font-black flex items-center justify-center">M</span>
                                <span>Men's Sleeveless V-Neck Sweaters (Designs: 01, 02, 03)</span>
                            </h5>
                            <span class="text-[10px] font-bold text-slate-500">Regular Comfort Fit</span>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-xs text-center border-collapse bg-white rounded-xl overflow-hidden border border-slate-200 shadow-sm">
                                <thead class="bg-slate-900 text-white text-[11px] font-bold">
                                    <tr>
                                        <th class="p-2.5 text-left">Size</th>
                                        <th class="p-2.5">Chest / Bust</th>
                                        <th class="p-2.5">Body Length</th>
                                        <th class="p-2.5">Shoulder</th>
                                        <th class="p-2.5 text-left">Recommended Body Build</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-100 font-medium">
                                    <tr class="hover:bg-slate-50">
                                        <td class="p-2.5 text-left font-black text-orange-600">S</td>
                                        <td class="p-2.5 font-bold text-slate-800">38"</td>
                                        <td class="p-2.5">26"</td>
                                        <td class="p-2.5">15"</td>
                                        <td class="p-2.5 text-left text-slate-600">Slim / Lean Build</td>
                                    </tr>
                                    <tr class="hover:bg-slate-50">
                                        <td class="p-2.5 text-left font-black text-orange-600">M</td>
                                        <td class="p-2.5 font-bold text-slate-800">40"</td>
                                        <td class="p-2.5">27"</td>
                                        <td class="p-2.5">16"</td>
                                        <td class="p-2.5 text-left text-slate-600">Medium Build (Standard)</td>
                                    </tr>
                                    <tr class="hover:bg-slate-50">
                                        <td class="p-2.5 text-left font-black text-orange-600">L</td>
                                        <td class="p-2.5 font-bold text-slate-800">42"</td>
                                        <td class="p-2.5">28"</td>
                                        <td class="p-2.5">17"</td>
                                        <td class="p-2.5 text-left text-slate-600">Standard Adult Fit</td>
                                    </tr>
                                    <tr class="hover:bg-slate-50">
                                        <td class="p-2.5 text-left font-black text-orange-600">XL</td>
                                        <td class="p-2.5 font-bold text-slate-800">44"</td>
                                        <td class="p-2.5">29"</td>
                                        <td class="p-2.5">18"</td>
                                        <td class="p-2.5 text-left text-slate-600">Plus / Comfort Fit</td>
                                    </tr>
                                    <tr class="hover:bg-slate-50">
                                        <td class="p-2.5 text-left font-black text-orange-600">XXL</td>
                                        <td class="p-2.5 font-bold text-slate-800">46"</td>
                                        <td class="p-2.5">30"</td>
                                        <td class="p-2.5">19"</td>
                                        <td class="p-2.5 text-left text-slate-600">Extra Comfort / Loose Fit</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- 2. Women's Cardigans Grid -->
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <!-- 04: Short Cardigan -->
                        <div class="bg-purple-50/50 rounded-2xl p-4 border border-purple-200 space-y-2.5">
                            <div class="flex items-center justify-between">
                                <h5 class="text-xs font-black text-purple-950 flex items-center gap-1.5">
                                    <span class="w-5 h-5 rounded-md bg-purple-600 text-white text-[10px] font-black flex items-center justify-center">04</span>
                                    <span>Women's Short Cardigan (Grid Check)</span>
                                </h5>
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-xs text-center border-collapse bg-white rounded-xl overflow-hidden border border-purple-200 shadow-sm">
                                    <thead class="bg-purple-900 text-white text-[11px] font-bold">
                                        <tr>
                                            <th class="p-2 text-left">Size</th>
                                            <th class="p-2">Chest</th>
                                            <th class="p-2">Length</th>
                                            <th class="p-2">Sleeve</th>
                                            <th class="p-2">Shoulder</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-purple-100 font-medium">
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">XS</td><td class="p-2 font-bold">34"</td><td class="p-2">21"</td><td class="p-2">21.5"</td><td class="p-2">13.5"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">S</td><td class="p-2 font-bold">36"</td><td class="p-2">22"</td><td class="p-2">22"</td><td class="p-2">14"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">M</td><td class="p-2 font-bold">38"</td><td class="p-2">23"</td><td class="p-2">22.5"</td><td class="p-2">14.5"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">L</td><td class="p-2 font-bold">40"</td><td class="p-2">24"</td><td class="p-2">23"</td><td class="p-2">15"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">XL</td><td class="p-2 font-bold">42"</td><td class="p-2">25"</td><td class="p-2">23.5"</td><td class="p-2">15.5"</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <!-- 05: Semi Long Cardigan -->
                        <div class="bg-purple-50/50 rounded-2xl p-4 border border-purple-200 space-y-2.5">
                            <div class="flex items-center justify-between">
                                <h5 class="text-xs font-black text-purple-950 flex items-center gap-1.5">
                                    <span class="w-5 h-5 rounded-md bg-purple-600 text-white text-[10px] font-black flex items-center justify-center">05</span>
                                    <span>Women's Semi Long Cardigan (Solid Black)</span>
                                </h5>
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-xs text-center border-collapse bg-white rounded-xl overflow-hidden border border-purple-200 shadow-sm">
                                    <thead class="bg-purple-900 text-white text-[11px] font-bold">
                                        <tr>
                                            <th class="p-2 text-left">Size</th>
                                            <th class="p-2">Chest</th>
                                            <th class="p-2">Length</th>
                                            <th class="p-2">Sleeve</th>
                                            <th class="p-2">Shoulder</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-purple-100 font-medium">
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">S</td><td class="p-2 font-bold">36"</td><td class="p-2">30"</td><td class="p-2">22"</td><td class="p-2">14"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">M</td><td class="p-2 font-bold">38"</td><td class="p-2">31"</td><td class="p-2">22.5"</td><td class="p-2">14.5"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">L</td><td class="p-2 font-bold">40"</td><td class="p-2">32"</td><td class="p-2">23"</td><td class="p-2">15"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">XL</td><td class="p-2 font-bold">42"</td><td class="p-2">33"</td><td class="p-2">23.5"</td><td class="p-2">15.5"</td></tr>
                                        <tr class="hover:bg-purple-50/50"><td class="p-2 text-left font-black text-purple-700">XXL</td><td class="p-2 font-bold">44"</td><td class="p-2">34"</td><td class="p-2">24"</td><td class="p-2">16"</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- ============================================== -->
    <!-- ADMIN PORTAL MODAL                             -->
    <!-- ============================================== -->
    <div id="admin-modal" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-3 sm:p-6" onclick="closeAdminModal()">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-5xl w-full max-h-[92vh] flex flex-col overflow-hidden" onclick="event.stopPropagation()">
            <!-- Admin Modal Header -->
            <div class="px-5 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-900 text-white">
                <div class="flex items-center gap-2.5">
                    <div class="w-8 h-8 rounded-xl bg-orange-600 text-white flex items-center justify-center font-bold text-sm shadow-sm"><i class="fa-solid fa-chart-pie"></i></div>
                    <div>
                        <h3 class="font-bold text-sm sm:text-base">Head Office Admin & Monitoring Portal</h3>
                        <p class="text-[10px] sm:text-xs text-slate-400">National Progress, Size Matrices, Cloud Sync & Live Master Excel</p>
                    </div>
                </div>
                <button onclick="closeAdminModal()" class="w-8 h-8 rounded-full bg-slate-800 border border-slate-700 text-slate-400 hover:text-white flex items-center justify-center transition"><i class="fa-solid fa-xmark text-sm"></i></button>
            </div>

            <!-- Admin Body -->
            <div class="p-4 sm:p-6 overflow-y-auto custom-scrollbar space-y-6">

                <!-- 1. National KPIs Summary -->
                <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4">
                    <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3.5 space-y-1">
                        <span class="text-[10px] font-bold text-slate-400 uppercase">Total Territories</span>
                        <div id="admin-kpi-total-terr" class="text-xl sm:text-2xl font-black text-slate-900">1,856</div>
                        <div class="text-[10px] text-slate-500 font-medium">35 Zones • 252 Regions</div>
                    </div>
                    <div class="bg-emerald-50 border border-emerald-200 rounded-2xl p-3.5 space-y-1">
                        <span class="text-[10px] font-bold text-emerald-700 uppercase">Completed</span>
                        <div id="admin-kpi-completed-terr" class="text-xl sm:text-2xl font-black text-emerald-700">0</div>
                        <div id="admin-kpi-completed-pct" class="text-[10px] text-emerald-600 font-bold">0% of National Goal</div>
                    </div>
                    <div class="bg-amber-50 border border-amber-200 rounded-2xl p-3.5 space-y-1">
                        <span class="text-[10px] font-bold text-amber-700 uppercase">In Progress</span>
                        <div id="admin-kpi-inprogress-terr" class="text-xl sm:text-2xl font-black text-amber-700">0</div>
                        <div class="text-[10px] text-amber-600 font-medium">Partially Filled</div>
                    </div>
                    <div class="bg-slate-100 border border-slate-300 rounded-2xl p-3.5 space-y-1">
                        <span class="text-[10px] font-bold text-slate-500 uppercase">Pending / Not Started</span>
                        <div id="admin-kpi-notstarted-terr" class="text-xl sm:text-2xl font-black text-slate-600">1,856</div>
                        <div class="text-[10px] text-slate-400 font-medium">Awaiting Submission</div>
                    </div>
                </div>

                <!-- 2. Master Action Buttons -->
                <div class="flex items-center gap-3 flex-wrap">
                    <button onclick="exportMasterExcelFromAdmin()" class="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white font-black text-xs sm:text-sm rounded-xl flex items-center gap-2 shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-file-excel text-base"></i>
                        <span>Export Live Master Excel</span>
                    </button>
                    <button onclick="pullCloudData(true)" class="px-4 py-2.5 bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-rotate"></i>
                        <span>Pull Cloud Data</span>
                    </button>
                    <button onclick="deleteAllCampaignData()" class="px-4 py-2.5 bg-rose-600 hover:bg-rose-500 text-white font-bold text-xs rounded-xl flex items-center gap-1.5 shadow-sm transition active:scale-95 ml-auto">
                        <i class="fa-solid fa-trash-can"></i>
                        <span>Delete All Data</span>
                    </button>
                </div>

                <!-- 3. Sweater & Size Production Matrix -->
                <div class="space-y-3">
                    <div class="flex items-center justify-between">
                        <h4 class="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-wide flex items-center gap-2">
                            <i class="fa-solid fa-boxes-stacked text-orange-500"></i>
                            <span>Aggregated Sweater Production Requirement Matrix</span>
                        </h4>
                        <span class="text-[10px] font-bold text-slate-500">Live counts across all 1,856 territories</span>
                    </div>
                    <div class="overflow-x-auto bg-white rounded-2xl border border-slate-200 shadow-sm">
                        <table class="w-full text-xs text-center border-collapse">
                            <thead class="bg-slate-900 text-white text-[11px] font-bold">
                                <tr>
                                    <th class="p-2.5 text-left">Item Code & Description</th>
                                    <th class="p-2.5">XS</th>
                                    <th class="p-2.5">S</th>
                                    <th class="p-2.5">M</th>
                                    <th class="p-2.5">L</th>
                                    <th class="p-2.5">XL</th>
                                    <th class="p-2.5">XXL</th>
                                    <th class="p-2.5 font-black text-orange-400">Total</th>
                                </tr>
                            </thead>
                            <tbody id="admin-production-matrix-body" class="divide-y divide-slate-100 font-semibold">
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- 4. Zone Progress Grid -->
                <div class="space-y-3">
                    <h4 class="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-wide flex items-center gap-2">
                        <i class="fa-solid fa-chart-simple text-orange-500"></i>
                        <span>Zone-wise Progress (35 Zones)</span>
                    </h4>
                    <div id="admin-zone-progress-grid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                    </div>
                </div>

                <!-- 5. Region-by-Region Table -->
                <div class="space-y-3">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <h4 class="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-wide flex items-center gap-2">
                            <i class="fa-solid fa-list-check text-orange-500"></i>
                            <span>All Regions Breakdown (252 Regions)</span>
                        </h4>
                        <input type="text" id="admin-region-search" oninput="renderAdminRegionsTable(this.value)" placeholder="Search Region, Head or SAP Code..." class="bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-800 w-full sm:w-64">
                    </div>
                    <div class="overflow-x-auto bg-white rounded-2xl border border-slate-200 shadow-sm max-h-80 custom-scrollbar">
                        <table class="w-full text-xs text-left border-collapse">
                            <thead class="bg-slate-100 text-slate-700 text-[11px] font-bold sticky top-0 z-10 border-b border-slate-200">
                                <tr>
                                    <th class="p-2.5">Region</th>
                                    <th class="p-2.5">Zone</th>
                                    <th class="p-2.5">Regional Head</th>
                                    <th class="p-2.5">Progress</th>
                                    <th class="p-2.5 text-right">Actions</th>
                                </tr>
                            </thead>
                            <tbody id="admin-regions-table-body" class="divide-y divide-slate-100">
                            </tbody>
                        </table>
                    </div>
                </div>

            </div>
        </div>
    </div>

    <!-- LIGHTBOX IMAGE MODAL -->
    <div id="lightbox-modal" class="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md hidden flex items-center justify-center p-4" onclick="closeImageLightbox()">
        <div class="max-w-md w-full bg-white rounded-3xl overflow-hidden shadow-2xl p-4 space-y-3" onclick="event.stopPropagation()">
            <div class="flex items-center justify-between pb-2 border-b border-slate-100">
                <span id="lightbox-badge" class="px-2.5 py-0.5 rounded-full text-xs font-black bg-orange-50 text-orange-700">Code: 01</span>
                <button onclick="closeImageLightbox()" class="text-slate-400 hover:text-slate-600 text-sm"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="aspect-[3/4] bg-slate-100 rounded-2xl overflow-hidden border border-slate-200 flex items-center justify-center">
                <img id="lightbox-img" src="" class="w-full h-full object-cover" alt="Preview">
            </div>
            <div class="text-center">
                <h4 id="lightbox-title" class="font-black text-slate-900 text-sm">--</h4>
                <p id="lightbox-desc" class="text-xs text-slate-500">--</p>
            </div>
        </div>
    </div>

    <!-- TOAST NOTIFICATION CONTAINER -->
    <div id="toast-container" class="fixed bottom-4 right-4 z-50 space-y-2 pointer-events-none"></div>

    <!-- JAVASCRIPT LOGIC -->
    <script>
        const REGION_MAP = ###REGION_MAP###;
        const ALL_TERRITORIES = ###ALL_TERRITORIES###;
        const ZONES = ###ZONES###;
        const DEFAULT_CLOUD_URL = "###DEFAULT_CLOUD_URL###";

        const SWEATER_DETAILS = {
            "01": { code: "01", name: "Men's Sleeveless V-Neck Sweater", color: "Solid Ash / Grey Textured", gender: "Men's", sizes: "S, M, L, XL, XXL", img: "###B64_01###", fallback_img: "Image/01 (Men).jpeg" },
            "02": { code: "02", name: "Men's Sleeveless V-Neck Sweater", color: "Solid Navy Blue Textured", gender: "Men's", sizes: "S, M, L, XL, XXL", img: "###B64_02###", fallback_img: "Image/02 (Men).jpeg" },
            "03": { code: "03", name: "Men's Sleeveless V-Neck Sweater", color: "Off-White / Cream Checkered", gender: "Men's", sizes: "S, M, L, XL, XXL", img: "###B64_03###", fallback_img: "Image/03 (Men).jpeg" },
            "04": { code: "04", name: "Women's Short Cardigan", color: "White & Navy Grid Check", gender: "Women's", sizes: "XS, S, M, L, XL", img: "###B64_04###", fallback_img: "Image/04 (Female).jpeg" },
            "05": { code: "05", name: "Women's Semi Long Cardigan", color: "Solid Black with Border Trim", gender: "Women's", sizes: "S, M, L, XL, XXL", img: "###B64_05###", fallback_img: "Image/05 (Female).jpeg" }
        };

        const ALL_TERRITORIES_MAP = {};
        ALL_TERRITORIES.forEach(t => {
            ALL_TERRITORIES_MAP[String(t['SAP Territory Code']).trim()] = t;
        });

        let store = JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{}');
        let cloudApiUrl = localStorage.getItem('EXIUM_CLOUD_URL') || DEFAULT_CLOUD_URL;

        let currentRegionCode = null;
        let activeTerritoryIndex = 0;
        let isAdminLoggedIn = false;
        let autoSyncTimeout = null;

        window.addEventListener('DOMContentLoaded', async () => {
            populateZoneDropdown();
            startDhakaClock();
            
            // Auto preload cloud data in background
            pullCloudData(false);

            const savedSession = JSON.parse(localStorage.getItem('EXIUM_ACTIVE_SESSION') || 'null');
            if (savedSession && savedSession.region_code && REGION_MAP[savedSession.region_code]) {
                unlockRegion(savedSession.region_code, true);
                if (typeof savedSession.territory_idx === 'number') {
                    selectTerritoryTab(savedSession.territory_idx, false);
                }
            }
        });

        function startDhakaClock() {
            function update() {
                const el = document.getElementById('live-dhaka-clock');
                if (el) {
                    const now = new Date();
                    el.textContent = now.toLocaleTimeString('en-US', { timeZone: 'Asia/Dhaka', hour: '2-digit', minute: '2-digit', second: '2-digit' }) + " (Dhaka)";
                }
            }
            update();
            setInterval(update, 1000);
        }

        // ==========================================
        // REGION LOGIN & DROPDOWNS
        // ==========================================
        function populateZoneDropdown() {
            const sel = document.getElementById('select-zone');
            if (!sel) return;
            if (sel.options.length <= 1) {
                sel.innerHTML = '<option value="">-- Choose Your Zone (35 Zones) --</option>';
                ZONES.forEach(z => {
                    const opt = document.createElement('option');
                    opt.value = z;
                    opt.textContent = z;
                    sel.appendChild(opt);
                });
            }
        }

        function onZoneChanged() {
            const zone = document.getElementById('select-zone').value;
            const regSel = document.getElementById('select-region');
            const rhCard = document.getElementById('rh-info-card');
            const passCard = document.getElementById('password-section');
            const btnCard = document.getElementById('unlock-btn-container');

            regSel.innerHTML = '<option value="">-- Select Region in ' + (zone || 'Zone') + ' --</option>';
            rhCard.classList.add('hidden');
            passCard.classList.add('hidden');
            btnCard.classList.add('hidden');

            if (!zone) {
                regSel.disabled = true;
                return;
            }

            regSel.disabled = false;
            for (let code in REGION_MAP) {
                const r = REGION_MAP[code];
                if (r.zone === zone) {
                    const opt = document.createElement('option');
                    opt.value = code;
                    opt.textContent = `${r.region_name} (${code}) - ${r.regional_head}`;
                    regSel.appendChild(opt);
                }
            }
        }

        function onRegionChanged() {
            const code = document.getElementById('select-region').value;
            const rhCard = document.getElementById('rh-info-card');
            const passCard = document.getElementById('password-section');
            const btnCard = document.getElementById('unlock-btn-container');

            if (!code || !REGION_MAP[code]) {
                rhCard.classList.add('hidden');
                passCard.classList.add('hidden');
                btnCard.classList.add('hidden');
                return;
            }

            const r = REGION_MAP[code];
            document.getElementById('card-region-name').textContent = r.region_name;
            document.getElementById('card-rh-name').textContent = r.regional_head;
            document.getElementById('card-sap-badge').textContent = `SAP: ${r.sap_region_code}`;

            rhCard.classList.remove('hidden');
            passCard.classList.remove('hidden');
            btnCard.classList.remove('hidden');

            document.getElementById('region-password').value = '';
            document.getElementById('region-password').focus();
        }

        function handlePasswordKey(e) {
            if (e.key === 'Enter') {
                unlockRegion();
            }
        }

        function togglePasswordVisibility() {
            const passInput = document.getElementById('region-password');
            const icon = document.getElementById('pass-toggle-icon');
            if (passInput.type === 'password') {
                passInput.type = 'text';
                icon.className = 'fa-regular fa-eye-slash';
            } else {
                passInput.type = 'password';
                icon.className = 'fa-regular fa-eye';
            }
        }

        function unlockRegion(bypassCode = null, isRestoringSession = false) {
            const code = bypassCode || document.getElementById('select-region').value;
            const passInput = document.getElementById('region-password');
            const pass = passInput ? passInput.value.trim() : '';

            if (!bypassCode && pass !== code && pass !== 'Exium MUPS' && pass !== '123456') {
                alert('Invalid Password! Please enter the correct password.');
                return;
            }

            if (!REGION_MAP[code]) return;

            currentRegionCode = code;
            pullCloudData(false);
            const r = REGION_MAP[code];

            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify({
                region_code: code,
                territory_idx: isRestoringSession ? activeTerritoryIndex : 0
            }));

            document.getElementById('selection-view').classList.add('hidden');
            document.getElementById('workspace-view').classList.remove('hidden');

            document.getElementById('banner-region').textContent = `SAP: ${r.sap_region_code}`;
            document.getElementById('banner-zone').textContent = r.zone;
            document.getElementById('banner-rh').textContent = `Region: ${r.region_name} (${r.regional_head})`;

            populateDropdownOptions();
            renderTerritoryTabs();
            selectTerritoryTab(isRestoringSession ? activeTerritoryIndex : 0);
            updateRegionalProgressBadge();
        }

        function exitRegionWorkspace() {
            localStorage.removeItem('EXIUM_ACTIVE_SESSION');
            currentRegionCode = null;
            document.getElementById('workspace-view').classList.add('hidden');
            document.getElementById('selection-view').classList.remove('hidden');
        }

        // ==========================================
        // TERRITORY NAVIGATION & TABS
        // ==========================================
        function populateDropdownOptions() {
            const sweaterSelects = ['c1_m1_sweater', 'c1_m2_sweater', 'c1_m3_sweater', 'c1_m4_sweater', 'c2_d1_sweater', 'c2_d2_sweater', 'c2_d3_sweater', 'c2_d4_sweater'];
            sweaterSelects.forEach(id => {
                const el = document.getElementById(id);
                if (el && el.options.length <= 1) {
                    el.innerHTML = '<option value="">-- Select Design --</option>';
                    for (let k in SWEATER_DETAILS) {
                        const item = SWEATER_DETAILS[k];
                        const opt = document.createElement('option');
                        opt.value = `${item.code} - ${item.name} (${item.color.split(' ')[0]})`;
                        opt.textContent = `${item.code} - ${item.name} (${item.gender}, ${item.color})`;
                        el.appendChild(opt);
                    }
                }
            });

            const sizeSelects = ['c1_m1_size', 'c1_m2_size', 'c1_m3_size', 'c1_m4_size', 'c2_d1_size', 'c2_d2_size', 'c2_d3_size', 'c2_d4_size'];
            const standardSizes = ['XS', 'S', 'M', 'L', 'XL', 'XXL'];
            sizeSelects.forEach(id => {
                const el = document.getElementById(id);
                if (el && el.options.length <= 1) {
                    el.innerHTML = '<option value="">-- Size --</option>';
                    standardSizes.forEach(sz => {
                        const opt = document.createElement('option');
                        opt.value = sz;
                        opt.textContent = sz;
                        el.appendChild(opt);
                    });
                }
            });
        }

        function renderTerritoryTabs() {
            if (!currentRegionCode) return;
            const r = REGION_MAP[currentRegionCode];
            const desktopContainer = document.getElementById('desktop-territory-tabs');
            const mobileSelect = document.getElementById('mobile-territory-select');

            desktopContainer.innerHTML = '';
            mobileSelect.innerHTML = '';

            r.territories.forEach((t, idx) => {
                const terrCode = String(t.sap_territory_code);
                const d = store[terrCode] || {};
                const status = getTerritoryStatus(d);

                // Desktop Button
                const btn = document.createElement('button');
                btn.id = `terr-tab-${idx}`;
                btn.onclick = () => selectTerritoryTab(idx);
                btn.className = `px-3.5 py-2 rounded-xl text-xs font-black whitespace-nowrap flex items-center gap-2 border transition ${
                    idx === activeTerritoryIndex ? 'bg-slate-900 text-white border-slate-900 shadow-sm' : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                }`;

                const pill = document.createElement('span');
                pill.className = `w-2 h-2 rounded-full ${
                    status === 'Complete' ? 'bg-emerald-500' : (status === 'In Progress' ? 'bg-amber-400' : 'bg-slate-300')
                }`;
                btn.appendChild(pill);

                const textSpan = document.createElement('span');
                textSpan.textContent = t.territory_name;
                btn.appendChild(textSpan);

                desktopContainer.appendChild(btn);

                // Mobile Select Option
                const opt = document.createElement('option');
                opt.value = idx;
                opt.textContent = `${t.territory_name} (${status})`;
                mobileSelect.appendChild(opt);
            });
        }

        function selectTerritoryTab(idx, shouldScroll = true) {
            activeTerritoryIndex = idx;
            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[idx];
            const terrCode = String(t.sap_territory_code);
            const d = store[terrCode] || {};

            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify({
                region_code: currentRegionCode,
                territory_idx: idx
            }));

            document.getElementById('mobile-territory-select').value = idx;
            document.getElementById('current-territory-title').textContent = t.territory_name;
            document.getElementById('current-territory-code').textContent = `SAP Code: ${terrCode}`;

            const status = getTerritoryStatus(d);
            const statusBadge = document.getElementById('current-territory-status');
            statusBadge.textContent = status;
            statusBadge.className = `text-[10px] font-bold px-2 py-0.5 rounded-full ${
                status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-bold' :
                'bg-slate-800 text-slate-300'
            }`;

            // Load Campaign 1
            document.getElementById('c1_doc_name').value = d.c1_doc_name || '';
            document.getElementById('c1_doc_rpl').value = d.c1_doc_rpl || '';
            document.getElementById('c1_m1_sweater').value = d.c1_m1_sweater || '';
            document.getElementById('c1_m1_size').value = d.c1_m1_size || '';
            document.getElementById('c1_m2_sweater').value = d.c1_m2_sweater || '';
            document.getElementById('c1_m2_size').value = d.c1_m2_size || '';
            document.getElementById('c1_m3_sweater').value = d.c1_m3_sweater || '';
            document.getElementById('c1_m3_size').value = d.c1_m3_size || '';
            document.getElementById('c1_m4_sweater').value = d.c1_m4_sweater || '';
            document.getElementById('c1_m4_size').value = d.c1_m4_size || '';

            // Load Campaign 2
            document.getElementById('c2_d1_name').value = d.c2_d1_name || '';
            document.getElementById('c2_d1_rpl').value = d.c2_d1_rpl || '';
            document.getElementById('c2_d1_sweater').value = d.c2_d1_sweater || '';
            document.getElementById('c2_d1_size').value = d.c2_d1_size || '';

            document.getElementById('c2_d2_name').value = d.c2_d2_name || '';
            document.getElementById('c2_d2_rpl').value = d.c2_d2_rpl || '';
            document.getElementById('c2_d2_sweater').value = d.c2_d2_sweater || '';
            document.getElementById('c2_d2_size').value = d.c2_d2_size || '';

            document.getElementById('c2_d3_name').value = d.c2_d3_name || '';
            document.getElementById('c2_d3_rpl').value = d.c2_d3_rpl || '';
            document.getElementById('c2_d3_sweater').value = d.c2_d3_sweater || '';
            document.getElementById('c2_d3_size').value = d.c2_d3_size || '';

            document.getElementById('c2_d4_name').value = d.c2_d4_name || '';
            document.getElementById('c2_d4_rpl').value = d.c2_d4_rpl || '';
            document.getElementById('c2_d4_sweater').value = d.c2_d4_sweater || '';
            document.getElementById('c2_d4_size').value = d.c2_d4_size || '';

            // Update Thumbnails
            updateSlotThumbnail('c1_m1_sweater', 'c1_m1_img_thumb');
            updateSlotThumbnail('c1_m2_sweater', 'c1_m2_img_thumb');
            updateSlotThumbnail('c1_m3_sweater', 'c1_m3_img_thumb');
            updateSlotThumbnail('c1_m4_sweater', 'c1_m4_img_thumb');

            // Validate RPLs on active form
            validateAllRplFields();

            // Update Tab styles
            r.territories.forEach((_, i) => {
                const btn = document.getElementById(`terr-tab-${i}`);
                if (btn) {
                    btn.className = `px-3.5 py-2 rounded-xl text-xs font-black whitespace-nowrap flex items-center gap-2 border transition ${
                        i === idx ? 'bg-slate-900 text-white border-slate-900 shadow-sm' : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                    }`;
                }
            });

            if (shouldScroll) {
                const activeBtn = document.getElementById(`terr-tab-${idx}`);
                if (activeBtn) activeBtn.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
            }
        }

        function updateSlotThumbnail(selectId, thumbContainerId) {
            const sel = document.getElementById(selectId);
            const container = document.getElementById(thumbContainerId);
            if (!sel || !container) return;

            const val = sel.value;
            const img = container.querySelector('img');
            if (!val) {
                if (img) img.classList.add('hidden');
                return;
            }

            const code = val.slice(0, 2);
            if (SWEATER_DETAILS[code]) {
                if (img) {
                    img.src = SWEATER_DETAILS[code].img;
                    img.classList.remove('hidden');
                }
            }
        }

        function openSlotLightbox(selectId) {
            const sel = document.getElementById(selectId);
            if (!sel || !sel.value) return;
            const code = sel.value.slice(0, 2);
            openImageLightbox(code);
        }

        // ==========================================
        // DOCTOR RPL DUPLICATE DETECTION ENGINE
        // ==========================================
        function findDoctorRplDuplicate(rplVal, currentFieldId) {
            if (!rplVal) return null;
            rplVal = String(rplVal).replace(/[^0-9]/g, '').trim();
            if (rplVal.length !== 6) return null;

            if (!currentRegionCode) return null;
            const r = REGION_MAP[currentRegionCode];
            const currentTerritory = r.territories[activeTerritoryIndex];
            const currentTerrCode = String(currentTerritory.sap_territory_code).trim();

            // 1. Check other fields on the ACTIVE form in DOM
            const activeFields = [
                { id: 'c1_doc_rpl', nameId: 'c1_doc_name', label: 'Campaign 1: Gyne Core Doctor' },
                { id: 'c2_d1_rpl', nameId: 'c2_d1_name', label: 'Campaign 2: Doctor 1' },
                { id: 'c2_d2_rpl', nameId: 'c2_d2_name', label: 'Campaign 2: Doctor 2' },
                { id: 'c2_d3_rpl', nameId: 'c2_d3_name', label: 'Campaign 2: Doctor 3' },
                { id: 'c2_d4_rpl', nameId: 'c2_d4_name', label: 'Campaign 2: Doctor 4' }
            ];

            for (const f of activeFields) {
                if (f.id === currentFieldId) continue;
                const el = document.getElementById(f.id);
                if (el) {
                    const elVal = String(el.value || '').replace(/[^0-9]/g, '').trim();
                    if (elVal === rplVal) {
                        const docNameEl = document.getElementById(f.nameId);
                        const docName = docNameEl ? docNameEl.value.trim() : '';
                        return {
                            isDuplicate: true,
                            sameTerritory: true,
                            territoryCode: currentTerrCode,
                            territoryName: currentTerritory.territory_name,
                            regionName: r.region_name,
                            regionalHead: r.regional_head,
                            zone: r.zone,
                            campaignLabel: f.label,
                            doctorName: docName || '(Doctor name not entered yet)'
                        };
                    }
                }
            }

            // 2. Check all territories across the entire campaign store
            const currentStore = Object.assign({}, store, JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{}'));

            for (const terrCode in currentStore) {
                const d = currentStore[terrCode];
                if (!d) continue;

                const isCurrentTerr = (String(terrCode).trim() === currentTerrCode);

                const storedSlots = [
                    { rplKey: 'c1_doc_rpl', nameKey: 'c1_doc_name', label: 'Campaign 1: Gyne Core Doctor', fieldId: 'c1_doc_rpl' },
                    { rplKey: 'c2_d1_rpl', nameKey: 'c2_d1_name', label: 'Campaign 2: Doctor 1', fieldId: 'c2_d1_rpl' },
                    { rplKey: 'c2_d2_rpl', nameKey: 'c2_d2_name', label: 'Campaign 2: Doctor 2', fieldId: 'c2_d2_rpl' },
                    { rplKey: 'c2_d3_rpl', nameKey: 'c2_d3_name', label: 'Campaign 2: Doctor 3', fieldId: 'c2_d3_rpl' },
                    { rplKey: 'c2_d4_rpl', nameKey: 'c2_d4_name', label: 'Campaign 2: Doctor 4', fieldId: 'c2_d4_rpl' }
                ];

                for (const slot of storedSlots) {
                    if (isCurrentTerr && slot.fieldId === currentFieldId) continue;
                    if (isCurrentTerr) continue; // Checked active DOM above

                    const slotRpl = String(d[slot.rplKey] || '').replace(/[^0-9]/g, '').trim();
                    if (slotRpl === rplVal) {
                        const terrMeta = ALL_TERRITORIES_MAP[String(terrCode).trim()] || {};
                        const docName = String(d[slot.nameKey] || '').trim();

                        return {
                            isDuplicate: true,
                            sameTerritory: false,
                            territoryCode: terrCode,
                            territoryName: terrMeta.Territory || `Territory ${terrCode}`,
                            regionName: terrMeta.Region || '',
                            regionalHead: terrMeta['Regional Head'] || '',
                            zone: terrMeta.Zone || '',
                            campaignLabel: slot.label,
                            doctorName: docName || '(Doctor name not entered yet)'
                        };
                    }
                }
            }

            return null;
        }

        function validateRplField(fieldId, badgeId) {
            const inputEl = document.getElementById(fieldId);
            const badgeEl = document.getElementById(badgeId);
            const msgEl = document.getElementById(fieldId + '_dup_msg');
            if (!inputEl) return true;

            const val = String(inputEl.value || '').replace(/[^0-9]/g, '').trim();
            if (val.length === 0) {
                if (badgeEl) {
                    badgeEl.textContent = "6 digits";
                    badgeEl.className = "text-[9px] sm:text-[10px] font-bold text-slate-400";
                }
                inputEl.classList.remove('border-rose-500', 'ring-2', 'ring-rose-300', 'bg-rose-50/50', 'text-rose-900');
                if (msgEl) { msgEl.className = 'hidden'; msgEl.innerHTML = ''; }
                return true;
            }

            if (val.length < 6) {
                if (badgeEl) {
                    badgeEl.textContent = `${val.length}/6 digits`;
                    badgeEl.className = "text-[9px] sm:text-[10px] font-bold text-amber-500";
                }
                inputEl.classList.remove('border-rose-500', 'ring-2', 'ring-rose-300', 'bg-rose-50/50', 'text-rose-900');
                if (msgEl) { msgEl.className = 'hidden'; msgEl.innerHTML = ''; }
                return true;
            }

            // 6 digits entered: Validate Duplicate
            const dup = findDoctorRplDuplicate(val, fieldId);
            if (dup && dup.isDuplicate) {
                if (badgeEl) {
                    badgeEl.textContent = "❌ ALREADY USED";
                    badgeEl.className = "text-[9px] sm:text-[10px] font-black text-rose-600 animate-pulse";
                }
                inputEl.classList.add('border-rose-500', 'ring-2', 'ring-rose-300', 'bg-rose-50/50', 'text-rose-900');
                if (msgEl) {
                    msgEl.className = 'mt-1.5 p-2.5 bg-rose-50 border border-rose-300 rounded-xl text-rose-900 text-[11px] shadow-sm space-y-1.5';
                    msgEl.innerHTML = `
                        <div class="flex items-center gap-1.5 font-black text-rose-700 text-xs">
                            <i class="fa-solid fa-circle-exclamation text-rose-600"></i>
                            <span>Doctor RPL ID (${val}) is ALREADY USED!</span>
                        </div>
                        <div class="bg-white/95 p-2 rounded-lg border border-rose-200 text-[10px] text-slate-700 space-y-0.5 leading-snug">
                            <div>📍 <strong>Territory:</strong> ${dup.territoryName} (Code: ${dup.territoryCode})</div>
                            <div>👤 <strong>Doctor:</strong> <span class="font-bold text-slate-900">${dup.doctorName}</span> (${dup.campaignLabel})</div>
                            <div>🌐 <strong>Region:</strong> ${dup.regionName} (${dup.regionalHead})</div>
                        </div>
                    `;
                }
                return false;
            } else {
                if (badgeEl) {
                    badgeEl.textContent = "✓ Valid";
                    badgeEl.className = "text-[9px] sm:text-[10px] font-bold text-emerald-600";
                }
                inputEl.classList.remove('border-rose-500', 'ring-2', 'ring-rose-300', 'bg-rose-50/50', 'text-rose-900');
                if (msgEl) { msgEl.className = 'hidden'; msgEl.innerHTML = ''; }
                return true;
            }
        }

        function validateAllRplFields() {
            let allValid = true;
            const rplPairs = [
                ['c1_doc_rpl', 'c1_doc_rpl_badge'],
                ['c2_d1_rpl', 'c2_d1_rpl_badge'],
                ['c2_d2_rpl', 'c2_d2_rpl_badge'],
                ['c2_d3_rpl', 'c2_d3_rpl_badge'],
                ['c2_d4_rpl', 'c2_d4_rpl_badge']
            ];

            rplPairs.forEach(([fieldId, badgeId]) => {
                const isValid = validateRplField(fieldId, badgeId);
                if (!isValid) allValid = false;
            });

            return allValid;
        }

        function onRplInput(inputEl, badgeId) {
            inputEl.value = inputEl.value.replace(/[^0-9]/g, '').slice(0, 6);
            validateAllRplFields();
            onDataChanged();
        }

        // ==========================================
        // DATA HANDLING & SAVE VALIDATION
        // ==========================================
        function onDataChanged() {
            if (!currentRegionCode) return;

            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[activeTerritoryIndex];
            const terrCode = String(t.sap_territory_code);

            const terrData = {
                c1_doc_name: document.getElementById('c1_doc_name').value.trim(),
                c1_doc_rpl: document.getElementById('c1_doc_rpl').value.trim(),
                c1_m1_sweater: document.getElementById('c1_m1_sweater').value,
                c1_m1_size: document.getElementById('c1_m1_size').value,
                c1_m2_sweater: document.getElementById('c1_m2_sweater').value,
                c1_m2_size: document.getElementById('c1_m2_size').value,
                c1_m3_sweater: document.getElementById('c1_m3_sweater').value,
                c1_m3_size: document.getElementById('c1_m3_size').value,
                c1_m4_sweater: document.getElementById('c1_m4_sweater').value,
                c1_m4_size: document.getElementById('c1_m4_size').value,

                c2_d1_name: document.getElementById('c2_d1_name').value.trim(),
                c2_d1_rpl: document.getElementById('c2_d1_rpl').value.trim(),
                c2_d1_sweater: document.getElementById('c2_d1_sweater').value,
                c2_d1_size: document.getElementById('c2_d1_size').value,

                c2_d2_name: document.getElementById('c2_d2_name').value.trim(),
                c2_d2_rpl: document.getElementById('c2_d2_rpl').value.trim(),
                c2_d2_sweater: document.getElementById('c2_d2_sweater').value,
                c2_d2_size: document.getElementById('c2_d2_size').value,

                c2_d3_name: document.getElementById('c2_d3_name').value.trim(),
                c2_d3_rpl: document.getElementById('c2_d3_rpl').value.trim(),
                c2_d3_sweater: document.getElementById('c2_d3_sweater').value,
                c2_d3_size: document.getElementById('c2_d3_size').value,

                c2_d4_name: document.getElementById('c2_d4_name').value.trim(),
                c2_d4_rpl: document.getElementById('c2_d4_rpl').value.trim(),
                c2_d4_sweater: document.getElementById('c2_d4_sweater').value,
                c2_d4_size: document.getElementById('c2_d4_size').value
            };

            store[terrCode] = terrData;
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));

            const status = getTerritoryStatus(terrData);
            const statusBadge = document.getElementById('current-territory-status');
            statusBadge.textContent = status;
            statusBadge.className = `text-[10px] font-bold px-2 py-0.5 rounded-full ${
                status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-bold' :
                'bg-slate-800 text-slate-300'
            }`;

            renderTerritoryTabs();
            updateRegionalProgressBadge();

            // Auto-sync debounce to cloud (if all RPLs valid)
            if (validateAllRplFields()) {
                if (autoSyncTimeout) clearTimeout(autoSyncTimeout);
                autoSyncTimeout = setTimeout(() => {
                    syncTerritoryToCloud(terrCode, terrData);
                }, 1200);
            }
        }

        function getTerritoryStatus(d) {
            if (!d) return 'Not Started';

            const isC1Done = Boolean(d.c1_doc_name && d.c1_doc_rpl && String(d.c1_doc_rpl).length === 6 && d.c1_m1_sweater && d.c1_m1_size && d.c1_m2_sweater && d.c1_m2_size && d.c1_m3_sweater && d.c1_m3_size && d.c1_m4_sweater && d.c1_m4_size);
            const isC2Done = Boolean(d.c2_d1_name && d.c2_d1_rpl && String(d.c2_d1_rpl).length === 6 && d.c2_d1_sweater && d.c2_d1_size && d.c2_d2_name && d.c2_d2_rpl && String(d.c2_d2_rpl).length === 6 && d.c2_d2_sweater && d.c2_d2_size && d.c2_d3_name && d.c2_d3_rpl && String(d.c2_d3_rpl).length === 6 && d.c2_d3_sweater && d.c2_d3_size && d.c2_d4_name && d.c2_d4_rpl && String(d.c2_d4_rpl).length === 6 && d.c2_d4_sweater && d.c2_d4_size);

            if (isC1Done && isC2Done) return 'Complete';

            const hasAny = Boolean(d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater || d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater || d.c2_d2_name || d.c2_d3_name || d.c2_d4_name);
            return hasAny ? 'In Progress' : 'Not Started';
        }

        function updateRegionalProgressBadge() {
            if (!currentRegionCode) return;
            const r = REGION_MAP[currentRegionCode];
            let completed = 0;

            r.territories.forEach(t => {
                const d = store[String(t.sap_territory_code)] || {};
                if (getTerritoryStatus(d) === 'Complete') completed++;
            });

            const total = r.territories.length;
            const pct = total > 0 ? Math.round((completed / total) * 100) : 0;

            document.getElementById('region-progress-text').textContent = `${completed} / ${total} Complete`;
            document.getElementById('region-progress-pct').textContent = `${pct}%`;
        }

        // VALIDATE TERRITORY COMPLETENESS BEFORE SAVING (ALL-OR-NOTHING RULE)
        function validateTerritoryForSave() {
            // 1. Check for duplicate Doctor RPL IDs
            const isRplValid = validateAllRplFields();
            if (!isRplValid) {
                return {
                    valid: false,
                    message: "❌ Duplicate Doctor RPL ID detected! The same RPL ID cannot be assigned to multiple doctors. Please correct the highlighted red field."
                };
            }

            // 2. Validate Campaign 1: Gyne Core Doctor Development (Family Package)
            const c1_name = document.getElementById('c1_doc_name').value.trim();
            const c1_rpl = document.getElementById('c1_doc_rpl').value.trim();
            const c1_m1_sw = document.getElementById('c1_m1_sweater').value;
            const c1_m1_sz = document.getElementById('c1_m1_size').value;
            const c1_m2_sw = document.getElementById('c1_m2_sweater').value;
            const c1_m2_sz = document.getElementById('c1_m2_size').value;
            const c1_m3_sw = document.getElementById('c1_m3_sweater').value;
            const c1_m3_sz = document.getElementById('c1_m3_size').value;
            const c1_m4_sw = document.getElementById('c1_m4_sweater').value;
            const c1_m4_sz = document.getElementById('c1_m4_size').value;

            const c1_hasAny = Boolean(c1_name || c1_rpl || c1_m1_sw || c1_m1_sz || c1_m2_sw || c1_m2_sz || c1_m3_sw || c1_m3_sz || c1_m4_sw || c1_m4_sz);

            if (c1_hasAny) {
                if (!c1_name) {
                    document.getElementById('c1_doc_name').focus();
                    return { valid: false, message: "⚠️ Campaign 1: Please enter Doctor Name (or clear all fields if skipping)." };
                }
                if (!c1_rpl || c1_rpl.length !== 6) {
                    document.getElementById('c1_doc_rpl').focus();
                    return { valid: false, message: "⚠️ Campaign 1: Doctor RPL ID must be exactly 6 digits." };
                }
                if (!c1_m1_sw || !c1_m1_sz) {
                    return { valid: false, message: "⚠️ Campaign 1: Please select both Sweater Design and Size for Sweater 1." };
                }
                if (!c1_m2_sw || !c1_m2_sz) {
                    return { valid: false, message: "⚠️ Campaign 1: Please select both Sweater Design and Size for Sweater 2." };
                }
                if (!c1_m3_sw || !c1_m3_sz) {
                    return { valid: false, message: "⚠️ Campaign 1: Please select both Sweater Design and Size for Sweater 3." };
                }
                if (!c1_m4_sw || !c1_m4_sz) {
                    return { valid: false, message: "⚠️ Campaign 1: Please select both Sweater Design and Size for Sweater 4." };
                }
            }

            // 3. Validate Campaign 2: Core Doctor Maximization (Doctors 1 to 4)
            for (let i = 1; i <= 4; i++) {
                const d_name = document.getElementById(`c2_d${i}_name`).value.trim();
                const d_rpl = document.getElementById(`c2_d${i}_rpl`).value.trim();
                const d_sw = document.getElementById(`c2_d${i}_sweater`).value;
                const d_sz = document.getElementById(`c2_d${i}_size`).value;

                const hasAny = Boolean(d_name || d_rpl || d_sw || d_sz);

                if (hasAny) {
                    if (!d_name) {
                        document.getElementById(`c2_d${i}_name`).focus();
                        return { valid: false, message: `⚠️ Campaign 2 (Doctor ${i}): Please enter Doctor Name (or clear all 4 fields if skipping).` };
                    }
                    if (!d_rpl || d_rpl.length !== 6) {
                        document.getElementById(`c2_d${i}_rpl`).focus();
                        return { valid: false, message: `⚠️ Campaign 2 (Doctor ${i}): Doctor RPL ID must be exactly 6 digits.` };
                    }
                    if (!d_sw) {
                        document.getElementById(`c2_d${i}_sweater`).focus();
                        return { valid: false, message: `⚠️ Campaign 2 (Doctor ${i}): Please select a Sweater design.` };
                    }
                    if (!d_sz) {
                        document.getElementById(`c2_d${i}_size`).focus();
                        return { valid: false, message: `⚠️ Campaign 2 (Doctor ${i}): Please select a Sweater Size.` };
                    }
                }
            }

            return { valid: true };
        }

        async function saveCurrentTerritoryClick() {
            const check = validateTerritoryForSave();
            if (!check.valid) {
                alert(check.message);
                showToast(check.message);
                return;
            }

            onDataChanged();
            
            if (currentRegionCode) {
                const r = REGION_MAP[currentRegionCode];
                const t = r.territories[activeTerritoryIndex];
                const terrCode = String(t.sap_territory_code);
                await syncTerritoryToCloud(terrCode, store[terrCode]);
            }
            
            showToast("✅ Territory saved & synced to Google Sheet!");
        }

        // ==========================================
        // GOOGLE SHEET CLOUD INTEGRATION
        // ==========================================
        async function syncTerritoryToCloud(terrCode, terrData) {
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (!url || !terrData || !terrCode) return;

            try {
                const payload = {
                    action: "save_territory",
                    sap_territory_code: String(terrCode).trim(),
                    data: terrData
                };

                await fetch(url, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify(payload)
                });
            } catch (err) {
                console.warn("[Cloud Sync Error]:", err);
            }
        }

        function fetchCloudDataJsonp(url) {
            return new Promise((resolve, reject) => {
                const cbName = 'gas_cb_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
                const script = document.createElement('script');
                const sep = url.includes('?') ? '&' : '?';
                script.src = `${url}${sep}action=get_all&callback=${cbName}&_t=${Date.now()}`;
                
                let isResolved = false;
                window[cbName] = function(data) {
                    isResolved = true;
                    delete window[cbName];
                    if (script.parentNode) script.parentNode.removeChild(script);
                    resolve(data);
                };
                script.onerror = function(err) {
                    if (!isResolved) {
                        delete window[cbName];
                        if (script.parentNode) script.parentNode.removeChild(script);
                        reject(err);
                    }
                };
                document.head.appendChild(script);

                setTimeout(() => {
                    if (!isResolved) {
                        delete window[cbName];
                        if (script.parentNode) script.parentNode.removeChild(script);
                        reject(new Error("Timeout pulling from Google Sheet"));
                    }
                }, 15000);
            });
        }

        async function pullCloudData(showFeedback = false) {
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (!url) return null;

            if (showFeedback) showToast("🔄 Fetching latest live data from Google Sheet...");

            let json = null;

            // 1. Try direct fetch
            try {
                const sep = url.includes('?') ? '&' : '?';
                const res = await fetch(`${url}${sep}action=get_all&_t=${Date.now()}`);
                if (res.ok) {
                    json = await res.json();
                }
            } catch (fetchErr) {
                console.warn("[Cloud Fetch Failed, Trying JSONP fallback]:", fetchErr);
            }

            // 2. Fallback to JSONP
            if (!json || json.status !== 'success' || !json.store) {
                try {
                    json = await fetchCloudDataJsonp(url);
                } catch (jsonpErr) {
                    console.warn("[JSONP Pull Failed]:", jsonpErr);
                }
            }

            if (json && json.status === 'success' && json.store) {
                let populatedCount = 0;
                for (let k in json.store) {
                    const item = json.store[k];
                    if (item && typeof item === 'object') {
                        store[k] = Object.assign(store[k] || {}, item);
                        if (item.c1_doc_name || item.c1_doc_rpl || item.c1_m1_sweater || item.c2_d1_name || item.c2_d1_rpl || item.c2_d1_sweater) {
                            populatedCount++;
                        }
                    }
                }
                localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));
                validateAllRplFields();

                if (isAdminLoggedIn) {
                    renderAdminKpisAndSummaries();
                    renderAdminZoneProgress();
                    renderAdminProductionMatrix();
                    renderAdminRegionsTable(document.getElementById('admin-region-search')?.value || '');
                }
                if (showFeedback) {
                    showToast(`✅ Synced with Google Sheet! (${populatedCount} active entries updated)`);
                }
                return { success: true, count: populatedCount };
            } else {
                if (showFeedback) {
                    showToast("⚠️ Could not pull cloud data. Please verify your connection.");
                }
                return { success: false, count: 0 };
            }
        }

        // ==========================================
        // EXCEL EXPORT FUNCTIONS
        // ==========================================
        function exportCurrentRegionExcel() {
            if (!currentRegionCode || !REGION_MAP[currentRegionCode]) return;
            const r = REGION_MAP[currentRegionCode];
            const terrList = ALL_TERRITORIES.filter(t => String(t['SAP Region Code']) === String(currentRegionCode));
            const safeName = r.region_name.replace(/[^a-zA-Z0-9]/g, '_');
            generateAndDownloadExcel(terrList, `Exium_Sweater_${safeName}_Region_Export.xlsx`);
        }

        async function exportMasterExcelFromAdmin() {
            showToast("⏳ Fetching live data from Google Sheet...");
            const res = await pullCloudData(false);
            if (res && res.success) {
                showToast(`📊 Generating Master Excel (${res.count} active territories found)...`);
            } else {
                showToast("📊 Generating Excel with available data...");
            }
            generateAndDownloadExcel(ALL_TERRITORIES, "Exium_MUPS_Sweater_Campaign_2026_Master_Export.xlsx");
        }

        function generateAndDownloadExcel(territoryList, filename) {
            const currentStore = Object.assign({}, JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{}'), store);

            const c1Rows = [];
            const c2Rows = [];

            territoryList.forEach(t => {
                const terrCode = String(t['SAP Territory Code']).trim();
                const d = currentStore[terrCode] || {};

                c1Rows.push({
                    "Zone": t.Zone,
                    "SAP Region Code": t['SAP Region Code'],
                    "Region": t.Region,
                    "Regional Head": t['Regional Head'],
                    "SAP Territory Code": t['SAP Territory Code'],
                    "Territory": t.Territory,
                    "Doctor Name": d.c1_doc_name || '',
                    "Doctor RPL ID": d.c1_doc_rpl || '',
                    "Sweater 1": d.c1_m1_sweater || '',
                    "Size 1": d.c1_m1_size || '',
                    "Sweater 2": d.c1_m2_sweater || '',
                    "Size 2": d.c1_m2_size || '',
                    "Sweater 3": d.c1_m3_sweater || '',
                    "Size 3": d.c1_m3_size || '',
                    "Sweater 4": d.c1_m4_sweater || '',
                    "Size 4": d.c1_m4_size || '',
                    "Status": (d.c1_doc_name && d.c1_doc_rpl && String(d.c1_doc_rpl).length === 6 && d.c1_m1_sweater && d.c1_m1_size && d.c1_m2_sweater && d.c1_m2_size && d.c1_m3_sweater && d.c1_m3_size && d.c1_m4_sweater && d.c1_m4_size) ? "Complete" : (d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater ? "In Progress" : "Not Started")
                });

                c2Rows.push({
                    "Zone": t.Zone,
                    "SAP Region Code": t['SAP Region Code'],
                    "Region": t.Region,
                    "Regional Head": t['Regional Head'],
                    "SAP Territory Code": t['SAP Territory Code'],
                    "Territory": t.Territory,
                    "Doctor 1 Name": d.c2_d1_name || '',
                    "Doctor 1 RPL ID": d.c2_d1_rpl || '',
                    "Sweater 1": d.c2_d1_sweater || '',
                    "Size 1": d.c2_d1_size || '',
                    "Doctor 2 Name": d.c2_d2_name || '',
                    "Doctor 2 RPL ID": d.c2_d2_rpl || '',
                    "Sweater 2": d.c2_d2_sweater || '',
                    "Size 2": d.c2_d2_size || '',
                    "Doctor 3 Name": d.c2_d3_name || '',
                    "Doctor 3 RPL ID": d.c2_d3_rpl || '',
                    "Sweater 3": d.c2_d3_sweater || '',
                    "Size 3": d.c2_d3_size || '',
                    "Doctor 4 Name": d.c2_d4_name || '',
                    "Doctor 4 RPL ID": d.c2_d4_rpl || '',
                    "Sweater 4": d.c2_d4_sweater || '',
                    "Size 4": d.c2_d4_size || '',
                    "Status": (d.c2_d1_name && d.c2_d1_rpl && String(d.c2_d1_rpl).length === 6 && d.c2_d1_sweater && d.c2_d1_size && d.c2_d2_name && d.c2_d2_rpl && String(d.c2_d2_rpl).length === 6 && d.c2_d2_sweater && d.c2_d2_size && d.c2_d3_name && d.c2_d3_rpl && String(d.c2_d3_rpl).length === 6 && d.c2_d3_sweater && d.c2_d3_size && d.c2_d4_name && d.c2_d4_rpl && String(d.c2_d4_rpl).length === 6 && d.c2_d4_sweater && d.c2_d4_size) ? "Complete" : (d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater || d.c2_d2_name || d.c2_d2_rpl || d.c2_d3_name || d.c2_d4_name ? "In Progress" : "Not Started")
                });
            });

            const wb = XLSX.utils.book_new();
            const ws1 = XLSX.utils.json_to_sheet(c1Rows);
            const ws2 = XLSX.utils.json_to_sheet(c2Rows);

            XLSX.utils.book_append_sheet(wb, ws1, "Gyne Core Doctor (Family)");
            XLSX.utils.book_append_sheet(wb, ws2, "Core Doctor Maximization");

            XLSX.writeFile(wb, filename);
            showToast(`✅ Excel file downloaded with all live data!`);
        }

        // ==========================================
        // ADMIN DASHBOARD RENDERING & ACTIONS
        // ==========================================
        function openAdminModal() {
            isAdminLoggedIn = true;
            document.getElementById('admin-modal').classList.remove('hidden');
            pullCloudData(false);
            renderAdminKpisAndSummaries();
            renderAdminZoneProgress();
            renderAdminProductionMatrix();
            renderAdminRegionsTable('');
        }

        function closeAdminModal() {
            isAdminLoggedIn = false;
            document.getElementById('admin-modal').classList.add('hidden');
        }

        function renderAdminKpisAndSummaries() {
            let completed = 0;
            let inProgress = 0;
            let notStarted = 0;

            ALL_TERRITORIES.forEach(t => {
                const code = String(t['SAP Territory Code']);
                const d = store[code] || {};
                const st = getTerritoryStatus(d);
                if (st === 'Complete') completed++;
                else if (st === 'In Progress') inProgress++;
                else notStarted++;
            });

            const total = ALL_TERRITORIES.length;
            const pct = Math.round((completed / total) * 100);

            document.getElementById('admin-kpi-total-terr').textContent = total.toLocaleString();
            document.getElementById('admin-kpi-completed-terr').textContent = completed.toLocaleString();
            document.getElementById('admin-kpi-completed-pct').textContent = `${pct}% of National Goal`;
            document.getElementById('admin-kpi-inprogress-terr').textContent = inProgress.toLocaleString();
            document.getElementById('admin-kpi-notstarted-terr').textContent = notStarted.toLocaleString();
        }

        function renderAdminZoneProgress() {
            const container = document.getElementById('admin-zone-progress-grid');
            if (!container) return;
            container.innerHTML = '';

            const zoneStats = {};
            ZONES.forEach(z => {
                zoneStats[z] = { total: 0, completed: 0 };
            });

            ALL_TERRITORIES.forEach(t => {
                const z = t.Zone;
                if (!zoneStats[z]) zoneStats[z] = { total: 0, completed: 0 };
                zoneStats[z].total++;
                const d = store[String(t['SAP Territory Code'])] || {};
                if (getTerritoryStatus(d) === 'Complete') zoneStats[z].completed++;
            });

            for (let z in zoneStats) {
                const s = zoneStats[z];
                const pct = s.total > 0 ? Math.round((s.completed / s.total) * 100) : 0;

                const card = document.createElement('div');
                card.className = 'bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2';
                card.innerHTML = `
                    <div class="flex items-center justify-between">
                        <strong class="text-xs font-bold text-slate-800">${z}</strong>
                        <span class="text-[10px] font-black ${pct === 100 ? 'text-emerald-700' : 'text-slate-500'}">${s.completed}/${s.total} (${pct}%)</span>
                    </div>
                    <div class="w-full h-2 bg-slate-200 rounded-full overflow-hidden">
                        <div class="h-full bg-gradient-to-r from-orange-500 to-amber-500 rounded-full" style="width: ${pct}%"></div>
                    </div>
                `;
                container.appendChild(card);
            }
        }

        function renderAdminProductionMatrix() {
            const tbody = document.getElementById('admin-production-matrix-body');
            if (!tbody) return;
            tbody.innerHTML = '';

            const matrix = {
                "01": { name: "Men's Sleeveless V-Neck (Grey)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "02": { name: "Men's Sleeveless V-Neck (Navy)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "03": { name: "Men's Sleeveless V-Neck (Cream)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "04": { name: "Women's Short Cardigan (Check)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "05": { name: "Women's Semi Long Cardigan (Black)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }
            };

            function tally(sw, sz) {
                if (!sw || !sz) return;
                const code = sw.slice(0, 2);
                if (matrix[code] && matrix[code][sz] !== undefined) {
                    matrix[code][sz]++;
                    matrix[code].total++;
                }
            }

            for (let k in store) {
                const d = store[k];
                if (!d) continue;
                tally(d.c1_m1_sweater, d.c1_m1_size);
                tally(d.c1_m2_sweater, d.c1_m2_size);
                tally(d.c1_m3_sweater, d.c1_m3_size);
                tally(d.c1_m4_sweater, d.c1_m4_size);
                tally(d.c2_d1_sweater, d.c2_d1_size);
                tally(d.c2_d2_sweater, d.c2_d2_size);
                tally(d.c2_d3_sweater, d.c2_d3_size);
                tally(d.c2_d4_sweater, d.c2_d4_size);
            }

            for (let code in matrix) {
                const row = matrix[code];
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="p-2.5 text-left font-bold text-slate-800"><span class="px-1.5 py-0.5 rounded bg-slate-100 text-slate-700 font-mono mr-1.5">${code}</span>${row.name}</td>
                    <td class="p-2.5 text-slate-700">${row.XS}</td>
                    <td class="p-2.5 text-slate-700">${row.S}</td>
                    <td class="p-2.5 text-slate-700">${row.M}</td>
                    <td class="p-2.5 text-slate-700">${row.L}</td>
                    <td class="p-2.5 text-slate-700">${row.XL}</td>
                    <td class="p-2.5 text-slate-700">${row.XXL}</td>
                    <td class="p-2.5 font-black text-orange-600 bg-orange-50/50">${row.total}</td>
                `;
                tbody.appendChild(tr);
            }
        }

        function renderAdminRegionsTable(searchQuery = '') {
            const tbody = document.getElementById('admin-regions-table-body');
            if (!tbody) return;
            tbody.innerHTML = '';

            const q = searchQuery.toLowerCase().trim();

            for (let code in REGION_MAP) {
                const r = REGION_MAP[code];
                if (q) {
                    const match = r.region_name.toLowerCase().includes(q) || r.regional_head.toLowerCase().includes(q) || r.zone.toLowerCase().includes(q) || code.includes(q);
                    if (!match) continue;
                }

                let completed = 0;
                r.territories.forEach(t => {
                    const d = store[String(t.sap_territory_code)] || {};
                    if (getTerritoryStatus(d) === 'Complete') completed++;
                });

                const total = r.territories.length;
                const pct = total > 0 ? Math.round((completed / total) * 100) : 0;

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="p-2.5 font-bold text-slate-900">${r.region_name} <span class="text-[10px] text-slate-400 font-normal">(${code})</span></td>
                    <td class="p-2.5 text-slate-600">${r.zone}</td>
                    <td class="p-2.5 text-slate-700 font-medium">${r.regional_head}</td>
                    <td class="p-2.5">
                        <span class="px-2 py-0.5 rounded-full text-[10px] font-black ${
                            pct === 100 ? 'bg-emerald-100 text-emerald-800' : (completed > 0 ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-500')
                        }">${completed}/${total} (${pct}%)</span>
                    </td>
                    <td class="p-2.5 text-right space-x-1.5">
                        <button onclick="unlockRegion('${code}')" class="px-2.5 py-1 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-[10px] font-bold">Open</button>
                        <button onclick="deleteSingleRegionData('${code}')" class="px-2 py-1 bg-rose-100 hover:bg-rose-200 text-rose-700 rounded-lg text-[10px] font-bold">Clear</button>
                    </td>
                `;
                tbody.appendChild(tr);
            }
        }

        async function deleteSingleRegionData(regCode) {
            if (!regCode || !REGION_MAP[regCode]) return;
            const r = REGION_MAP[regCode];
            if (!confirm(`Are you sure you want to clear all data for Region: ${r.region_name} (${regCode})?`)) return;

            r.territories.forEach(t => {
                delete store[String(t.sap_territory_code)];
            });
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));

            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (url) {
                try {
                    await fetch(url, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                        body: JSON.stringify({ action: "delete_region", sap_region_code: String(regCode).trim() })
                    });
                } catch (e) {
                    console.warn(e);
                }
            }

            renderAdminKpisAndSummaries();
            renderAdminZoneProgress();
            renderAdminProductionMatrix();
            renderAdminRegionsTable(document.getElementById('admin-region-search')?.value || '');
            showToast(`🗑️ Data for ${r.region_name} cleared from portal and Google Sheet.`);
        }

        async function deleteAllCampaignData() {
            const promptVal = prompt("⚠️ DANGER ZONE: This will wipe ALL entered data across all 1,856 territories in both the Web Portal and Google Sheet!\\n\\nType 'DELETE ALL' to confirm:");
            if (promptVal !== 'DELETE ALL') {
                alert("Action cancelled.");
                return;
            }

            store = {};
            localStorage.removeItem('EXIUM_SWEATER_STORE');
            localStorage.removeItem('EXIUM_ACTIVE_SESSION');

            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (url) {
                try {
                    await fetch(url, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                        body: JSON.stringify({ action: "reset_all" })
                    });
                } catch (e) {
                    console.warn(e);
                }
            }

            renderAdminKpisAndSummaries();
            renderAdminZoneProgress();
            renderAdminProductionMatrix();
            renderAdminRegionsTable('');
            showToast("🗑️ All campaign data reset across portal and Google Sheet!");
        }

        // ==========================================
        // MODALS & LIGHTBOX
        // ==========================================
        function openCatalogModal() {
            document.getElementById('catalog-modal').classList.remove('hidden');
        }

        function closeCatalogModal() {
            document.getElementById('catalog-modal').classList.add('hidden');
        }

        function openImageLightbox(code) {
            const item = SWEATER_DETAILS[code];
            if (!item) return;

            document.getElementById('lightbox-badge').textContent = `Design: ${item.code} (${item.gender})`;
            document.getElementById('lightbox-title').textContent = item.name;
            document.getElementById('lightbox-desc').textContent = `${item.color} • Available Sizes: ${item.sizes}`;
            document.getElementById('lightbox-img').src = item.img;

            document.getElementById('lightbox-modal').classList.remove('hidden');
        }

        function closeImageLightbox() {
            document.getElementById('lightbox-modal').classList.add('hidden');
        }

        function showToast(msg) {
            const container = document.getElementById('toast-container');
            if (!container) return;

            const toast = document.createElement('div');
            toast.className = 'bg-slate-900 text-white px-4 py-2.5 rounded-2xl shadow-xl text-xs font-bold flex items-center gap-2 border border-slate-700 animate-fade-in pointer-events-auto';
            toast.innerHTML = `<span>${msg}</span>`;

            container.appendChild(toast);
            setTimeout(() => {
                toast.style.opacity = '0';
                toast.style.transition = 'opacity 0.3s';
                setTimeout(() => toast.remove(), 300);
            }, 3200);
        }
    </script>
</body>
</html>
"""

# Embed base64 images and metadata
html_output = html_template.replace("###B64_LOGO_SQ###", b64_logo_sq)
html_output = html_output.replace("###B64_LOGO_BANNER###", b64_logo_banner)
html_output = html_output.replace("###B64_01###", b64_01)
html_output = html_output.replace("###B64_02###", b64_02)
html_output = html_output.replace("###B64_03###", b64_03)
html_output = html_output.replace("###B64_04###", b64_04)
html_output = html_output.replace("###B64_05###", b64_05)
html_output = html_output.replace("###REGION_MAP###", json.dumps(region_map))
html_output = html_output.replace("###ALL_TERRITORIES###", json.dumps(territories))
html_output = html_output.replace("###ZONES###", json.dumps(zones))
html_output = html_output.replace("###DEFAULT_CLOUD_URL###", default_cloud_url)

with open(r"G:\Exium\2026\4Q'26\Sweater\index.html", "w", encoding="utf-8") as f:
    f.write(html_output)

with open(r"G:\Exium\2026\4Q'26\Sweater\Sweater_Campaign_Portal.html", "w", encoding="utf-8") as f:
    f.write(html_output)

print("Successfully regenerated clean and complete index.html and Sweater_Campaign_Portal.html!")
