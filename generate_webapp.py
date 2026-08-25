import os
import pandas as pd
import base64
import json

base_dir = r"G:\Exium\2026\4Q'26\Sweater"
excel_file = os.path.join(base_dir, "FF list.xlsx")
image_dir = os.path.join(base_dir, "Image")
logo_path = os.path.join(base_dir, "Exium MUPS Logo.png")

print("Encoding images...")
def get_base64_img(path):
    if os.path.exists(path):
        with open(path, "rb") as f:
            ext = os.path.splitext(path)[1].lower().replace('.', '')
            if ext == 'jpg': ext = 'jpeg'
            return f"data:image/{ext};base64,{base64.b64encode(f.read()).decode('utf-8')}"
    return ""

b64_logo = get_base64_img(logo_path)
b64_01 = get_base64_img(os.path.join(image_dir, "01 (Men).jpeg"))
b64_02 = get_base64_img(os.path.join(image_dir, "02 (Men).jpeg"))
b64_03 = get_base64_img(os.path.join(image_dir, "03 (Men).jpeg"))
b64_04 = get_base64_img(os.path.join(image_dir, "04 (Female).jpeg"))
b64_05 = get_base64_img(os.path.join(image_dir, "05 (Female).jpeg"))

df = pd.read_excel(excel_file)
territories = df.to_dict(orient='records')

region_map = {}
zone_map = {}

for t in territories:
    z = str(t['Zone']).strip()
    reg_code = str(t['SAP Region Code']).strip()
    reg_name = str(t['Region']).strip()
    reg_head = str(t['Regional Head']).strip()
    terr_code = str(t['SAP Territory Code']).strip()
    terr_name = str(t['Territory']).strip()

    if z not in zone_map:
        zone_map[z] = []
    if reg_code not in zone_map[z]:
        zone_map[z].append(reg_code)

    if reg_code not in region_map:
        region_map[reg_code] = {
            'sap_region_code': reg_code,
            'region_name': reg_name,
            'regional_head': reg_head,
            'zone': z,
            'territories': []
        }
    region_map[reg_code]['territories'].append({
        'sap_territory_code': terr_code,
        'territory_name': terr_name
    })

zones = sorted(list(zone_map.keys()))
zone_options = ""
for z in zones:
    zone_options += f'<option value="{z}">{z}</option>\n'

DEFAULT_CLOUD_URL = "https://script.google.com/macros/s/AKfycbzEnDTtNiXEAyB5qHqrxLj1RbNytgOJAB_lKjw_VVVd1C8CiaeYU6iTROiJabkyX_-b/exec"

html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exium MUPS - Sweater Campaign 2026</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        body { font-family: 'Plus Jakarta Sans', sans-serif; }
        .font-mono { font-family: 'JetBrains Mono', monospace; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 9999px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 9999px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        .sweater-card-img { transition: transform 0.2s cubic-bezier(0.4, 0, 0.2, 1); }
        .sweater-card-img:hover { transform: scale(1.03); }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col antialiased">

    <!-- TOP HEADER -->
    <header class="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-2 sm:py-2.5">
            <div class="flex items-center justify-between gap-3">
                
                <!-- Left: Logo + Title -->
                <div class="flex items-center gap-2 sm:gap-2.5 min-w-0">
                    <img src="###B64_LOGO###" onerror="this.src='Exium MUPS Logo.png'" alt="Exium MUPS" class="h-7 sm:h-8 md:h-9 w-auto object-contain flex-shrink-0">
                    <div class="border-l-2 border-slate-300 pl-2 sm:pl-2.5 flex items-center gap-1.5 sm:gap-2 min-w-0">
                        <h1 class="text-sm sm:text-base md:text-lg font-black text-slate-900 tracking-tight leading-none whitespace-nowrap">Sweater for Doctors</h1>
                        <span class="text-[10px] sm:text-xs bg-orange-500 text-white font-black px-1.5 sm:px-2 py-0.5 rounded-full leading-none shadow-sm flex-shrink-0">4Q'26</span>
                    </div>
                </div>

                <!-- Right Alignment: Catalogue & Sizes + Admin -->
                <div class="flex items-center gap-2 flex-shrink-0">
                    <button onclick="openCatalogModal()" class="px-2.5 sm:px-3.5 py-1.5 bg-orange-50 hover:bg-orange-100 text-orange-800 border border-orange-200 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm active:scale-95 whitespace-nowrap">
                        <i class="fa-solid fa-vest text-orange-600"></i>
                        <span class="hidden xs:inline">Catalogue & Sizes</span>
                        <span class="xs:hidden">Catalog</span>
                    </button>
                    <div id="header-admin-btn-container">
                        <button onclick="openAdminModal()" class="px-2.5 sm:px-3.5 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm active:scale-95 whitespace-nowrap">
                            <i class="fa-solid fa-shield-halved text-orange-400"></i>
                            <span>Admin</span>
                        </button>
                    </div>
                </div>

            </div>
        </div>
    </header>

    <!-- MAIN WRAPPER -->
    <div class="flex-1 flex flex-col">

        <!-- 1. SELECTION / LOGIN VIEW -->
        <div id="selection-view" class="max-w-xl w-full mx-auto px-4 py-8 sm:py-12 flex-1 flex flex-col justify-center">
            <div class="bg-white rounded-3xl border border-slate-200 shadow-xl overflow-hidden">
                <div class="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white p-6 sm:p-8 text-center relative">
                    <div class="w-14 h-14 bg-white/10 rounded-2xl flex items-center justify-center mx-auto mb-3 border border-white/20 shadow-inner">
                        <i class="fa-solid fa-user-tie text-2xl text-orange-400"></i>
                    </div>
                    <h2 class="text-xl sm:text-2xl font-black tracking-tight">Regional Manager Login</h2>
                    <p class="text-slate-300 text-xs sm:text-sm mt-1">Select your Zone & Region to enter territory requisitions</p>
                </div>

                <div class="p-6 sm:p-8 space-y-5">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">1. Select Zone</label>
                        <select id="select-zone" onchange="onZoneChanged()" class="w-full bg-slate-50 border border-slate-300 rounded-2xl px-4 py-3 text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition">
                            <option value="">-- Choose Your Zone (35 Zones) --</option>
                            ###ZONE_OPTIONS###
                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">2. Select Region (Regional Head)</label>
                        <select id="select-region" onchange="onRegionChanged()" disabled class="w-full bg-slate-100 border border-slate-300 rounded-2xl px-4 py-3 text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition disabled:opacity-60 disabled:cursor-not-allowed">
                            <option value="">-- First Select Zone Above --</option>
                        </select>
                    </div>

                    <div id="password-container" class="hidden space-y-2 pt-2 border-t border-slate-100">
                        <div class="flex items-center justify-between">
                            <label class="text-xs font-bold text-slate-700 uppercase tracking-wider">3. Enter Security PIN / Password</label>
                            <span class="text-[10px] text-slate-400 font-medium">Default: 1234</span>
                        </div>
                        <div class="relative">
                            <input type="password" id="region-password" onkeyup="handlePasswordKey(event)" placeholder="Enter 4-digit PIN..." class="w-full bg-slate-50 border border-slate-300 rounded-2xl px-4 py-3 text-sm font-mono font-bold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none tracking-widest transition">
                            <button type="button" onclick="togglePasswordVisibility('region-password')" class="absolute right-3.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 p-1 text-sm"><i class="fa-solid fa-eye"></i></button>
                        </div>
                    </div>

                    <button id="btn-unlock-region" onclick="unlockRegion()" disabled class="w-full py-3.5 px-4 bg-orange-500 hover:bg-orange-600 disabled:bg-slate-200 text-white disabled:text-slate-400 rounded-2xl font-black text-sm tracking-wide shadow-md hover:shadow-lg transition flex items-center justify-center gap-2 disabled:cursor-not-allowed active:scale-[0.99]">
                        <i class="fa-solid fa-arrow-right-to-bracket"></i>
                        <span>Enter Workspace</span>
                    </button>

                    <div class="text-center pt-2">
                        <button type="button" onclick="openAdminModal()" class="text-xs font-bold text-slate-400 hover:text-slate-700 transition inline-flex items-center gap-1">
                            <i class="fa-solid fa-shield-halved text-orange-500"></i>
                            <span>Admin Portal & Live Reports</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>

        <!-- 2. WORKSPACE VIEW (REGIONAL MANAGER PORTAL) -->
        <main id="workspace-view" class="hidden flex-1 max-w-7xl w-full mx-auto px-3 sm:px-6 lg:px-8 py-4 sm:py-6 space-y-4 sm:space-y-6">
            
            <!-- Region Banner -->
            <div class="bg-white rounded-3xl p-4 sm:p-5 border border-slate-200 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                    <div class="flex items-center gap-2 mb-1">
                        <span id="banner-region" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200">SAP: 00000</span>
                        <span id="region-progress-badge" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">0/0 Done</span>
                    </div>
                    <h2 id="banner-rh" class="text-lg sm:text-2xl font-black text-slate-900 tracking-tight">Region: Region Name (Regional Head)</h2>
                </div>

                <div class="flex items-center gap-2 self-start sm:self-auto">
                    <button onclick="exportCurrentRegionExcel()" class="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-file-excel"></i>
                        <span>Export Region Excel</span>
                    </button>
                    <button onclick="exitRegionWorkspace()" class="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm transition active:scale-95">
                        <i class="fa-solid fa-right-from-bracket"></i>
                        <span>Exit</span>
                    </button>
                </div>
            </div>

            <!-- Workspace Layout (Sidebar + Main Form) -->
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-4 sm:gap-6">

                <!-- Left Sidebar: Territory List -->
                <div class="lg:col-span-3 space-y-3">
                    <div class="block lg:hidden bg-white border border-slate-200 rounded-2xl p-3 shadow-sm">
                        <label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">Switch Territory</label>
                        <select id="mobile-territory-select" onchange="selectTerritoryTab(parseInt(this.value))" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none">
                        </select>
                    </div>

                    <div class="hidden lg:flex flex-col bg-white border border-slate-200 rounded-3xl p-4 shadow-sm space-y-3">
                        <div class="flex items-center justify-between border-b border-slate-100 pb-2">
                            <h3 class="text-xs font-black uppercase tracking-wider text-slate-500">Territories</h3>
                            <span id="sidebar-count" class="text-xs font-bold text-slate-400">12 Total</span>
                        </div>
                        <div id="desktop-territory-list" class="space-y-1.5 max-h-[calc(100vh-280px)] overflow-y-auto custom-scrollbar pr-1">
                            <!-- Territory Buttons rendered via JS -->
                        </div>
                    </div>
                </div>

                <!-- Right Main Content: Form -->
                <div class="lg:col-span-9 space-y-4 sm:space-y-6">

                    <!-- Active Territory Header Card -->
                    <div id="active-territory-banner-card" class="bg-slate-900 text-white rounded-3xl p-4 sm:p-5 flex items-center justify-between shadow-md">
                        <div>
                            <span class="text-[10px] font-bold uppercase tracking-wider text-orange-400">Current Active Territory</span>
                            <div class="flex items-center gap-2 mt-0.5">
                                <h3 id="current-territory-title" class="text-base sm:text-xl font-black">Territory Name</h3>
                                <span id="current-territory-code" class="text-xs text-slate-400 font-mono">SAP: 00000</span>
                            </div>
                        </div>
                        <div>
                            <span id="current-territory-status" class="text-[10px] sm:text-xs font-black px-3 py-1 rounded-full bg-white/10 text-slate-200 border border-white/20">Not Started</span>
                        </div>
                    </div>

                    <!-- CAMPAIGN 1: GYNE CORE DOCTOR (FAMILY PACKAGE) -->
                    <div class="bg-white border-2 border-teal-500/60 rounded-3xl shadow-sm overflow-hidden">
                        <div class="bg-gradient-to-r from-teal-700 via-teal-800 to-emerald-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-7 h-7 sm:w-8 sm:h-8 rounded-xl bg-white text-teal-800 flex items-center justify-center font-black text-xs sm:text-sm flex-shrink-0 shadow-sm mt-0.5 sm:mt-0">1</div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Gyne Core Doctor Development (Family Package)</h4>
                                    <p class="text-[10px] sm:text-xs text-teal-100 mt-0.5 leading-tight">1 Doctor per Territory &bull; 3 Sweaters (Option to Add 4th Sweater)</p>
                                </div>
                            </div>
                            <div class="self-start sm:self-auto pl-9 sm:pl-0">
                                <span class="text-[10px] sm:text-xs font-black bg-teal-950/80 text-teal-200 border border-teal-400/40 px-2.5 py-0.5 rounded-full inline-block whitespace-nowrap shadow-sm">Family Package</span>
                            </div>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
                            <!-- Doctor Info -->
                            <div class="bg-teal-50/70 rounded-2xl p-3 sm:p-4 border border-teal-200">
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                    <div>
                                        <label class="block text-xs font-bold text-teal-950 mb-1">Doctor Name <span class="text-rose-500">*</span></label>
                                        <input type="text" id="c1_doc_name" oninput="onDataChanged()" placeholder="Enter Gynecologist / Doctor Name..." class="w-full bg-white border border-teal-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 focus:outline-none transition">
                                    </div>
                                    <div>
                                        <div class="flex items-center justify-between mb-1">
                                            <label class="text-xs font-bold text-teal-950">Doctor RPL ID (6 Digits) <span class="text-rose-500">*</span></label>
                                            <span id="c1_doc_rpl_badge" class="text-[10px] font-bold text-slate-400">6 digits</span>
                                        </div>
                                        <input type="text" inputmode="numeric" maxlength="6" id="c1_doc_rpl" oninput="onRplInput(this, 'c1_doc_rpl_badge')" placeholder="e.g. 104523" class="w-full bg-white border border-teal-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 font-mono font-bold placeholder-slate-400 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 focus:outline-none transition tracking-wider">
                                    </div>
                                </div>
                            </div>

                            <!-- 2x2 Grid: Option 1 & 2 on Top, Option 3 & 4 on Bottom -->
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4">
                                
                                <!-- 1. Sweater 1 (Doctor / Family Member) -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">1</span> Sweater 1 (Doctor / Family Member)</span>
                                        <span id="c1_m1_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m1_img_preview" onclick="zoomSlotImage('c1_m1_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                        <div class="flex-1 space-y-1.5 min-w-0">
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                <select id="c1_m1_sweater" onchange="onSweaterSelectChange('c1_m1', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Select Sweater --</option>
                                                    <option value="01 - Men's V-Neck (Grey)">01 - Men's V-Neck (Grey)</option>
                                                    <option value="02 - Men's V-Neck (Navy Blue)">02 - Men's V-Neck (Navy Blue)</option>
                                                    <option value="03 - Men's V-Neck (Cream Check)">03 - Men's V-Neck (Cream Check)</option>
                                                    <option value="04 - Women's Short Cardigan (Check)">04 - Women's Short Cardigan</option>
                                                    <option value="05 - Women's Semi Long Cardigan (Black)">05 - Women's Semi Long</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Size</label>
                                                <select id="c1_m1_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500"><option value="">-- Size --</option></select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- 2. Sweater 2 (Family Member) -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">2</span> Sweater 2 (Family Member)</span>
                                        <span id="c1_m2_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m2_img_preview" onclick="zoomSlotImage('c1_m2_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                        <div class="flex-1 space-y-1.5 min-w-0">
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                <select id="c1_m2_sweater" onchange="onSweaterSelectChange('c1_m2', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Select Sweater --</option>
                                                    <option value="01 - Men's V-Neck (Grey)">01 - Men's V-Neck (Grey)</option>
                                                    <option value="02 - Men's V-Neck (Navy Blue)">02 - Men's V-Neck (Navy Blue)</option>
                                                    <option value="03 - Men's V-Neck (Cream Check)">03 - Men's V-Neck (Cream Check)</option>
                                                    <option value="04 - Women's Short Cardigan (Check)">04 - Women's Short Cardigan</option>
                                                    <option value="05 - Women's Semi Long Cardigan (Black)">05 - Women's Semi Long</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Size</label>
                                                <select id="c1_m2_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500"><option value="">-- Size --</option></select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- 3. Sweater 3 (Family Member) -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">3</span> Sweater 3 (Family Member)</span>
                                        <span id="c1_m3_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m3_img_preview" onclick="zoomSlotImage('c1_m3_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                        <div class="flex-1 space-y-1.5 min-w-0">
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                <select id="c1_m3_sweater" onchange="onSweaterSelectChange('c1_m3', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Select Sweater --</option>
                                                    <option value="01 - Men's V-Neck (Grey)">01 - Men's V-Neck (Grey)</option>
                                                    <option value="02 - Men's V-Neck (Navy Blue)">02 - Men's V-Neck (Navy Blue)</option>
                                                    <option value="03 - Men's V-Neck (Cream Check)">03 - Men's V-Neck (Cream Check)</option>
                                                    <option value="04 - Women's Short Cardigan (Check)">04 - Women's Short Cardigan</option>
                                                    <option value="05 - Women's Semi Long Cardigan (Black)">05 - Women's Semi Long</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Size</label>
                                                <select id="c1_m3_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500"><option value="">-- Size --</option></select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- 4. Add Another Sweater Button Container -->
                                <div id="c1_add_m4_btn_container" onclick="showC1Sweater4(true)" class="border-2 border-dashed border-teal-300 hover:border-teal-500 bg-teal-50/40 hover:bg-teal-50 rounded-2xl p-4 flex flex-col items-center justify-center text-center cursor-pointer transition group gap-2 min-h-[140px]">
                                    <div class="w-10 h-10 rounded-full bg-teal-100 group-hover:bg-teal-600 text-teal-600 group-hover:text-white flex items-center justify-center transition shadow-sm font-black"><i class="fa-solid fa-plus text-base"></i></div>
                                    <div>
                                        <h5 class="text-xs font-black text-teal-950 group-hover:text-teal-700">Add Another Sweater</h5>
                                        <p class="text-[10px] text-slate-500">Sweater 4 (Family Member)</p>
                                    </div>
                                </div>

                                <!-- 4. Sweater 4 (Family Member) - Revealed on Click -->
                                <div id="c1_m4_container" class="hidden bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">4</span> Sweater 4 (Family Member)</span>
                                        <div class="flex items-center gap-1.5">
                                            <span id="c1_m4_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                            <button type="button" onclick="hideAndClearC1Sweater4()" title="Remove 4th Sweater" class="text-[10px] text-rose-500 hover:text-rose-700 font-bold px-1.5 py-0.5 rounded hover:bg-rose-50"><i class="fa-solid fa-xmark"></i></button>
                                        </div>
                                    </div>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m4_img_preview" onclick="zoomSlotImage('c1_m4_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                        <div class="flex-1 space-y-1.5 min-w-0">
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                <select id="c1_m4_sweater" onchange="onSweaterSelectChange('c1_m4', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Select Sweater --</option>
                                                    <option value="01 - Men's V-Neck (Grey)">01 - Men's V-Neck (Grey)</option>
                                                    <option value="02 - Men's V-Neck (Navy Blue)">02 - Men's V-Neck (Navy Blue)</option>
                                                    <option value="03 - Men's V-Neck (Cream Check)">03 - Men's V-Neck (Cream Check)</option>
                                                    <option value="04 - Women's Short Cardigan (Check)">04 - Women's Short Cardigan</option>
                                                    <option value="05 - Women's Semi Long Cardigan (Black)">05 - Women's Semi Long</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Size</label>
                                                <select id="c1_m4_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500"><option value="">-- Size --</option></select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                            </div>
                        </div>
                    </div>

                    <!-- CAMPAIGN 2: CORE DOCTOR MAXIMIZATION (3 DOCTORS) -->
                    <div class="bg-white border-2 border-purple-500/60 rounded-3xl shadow-sm overflow-hidden">
                        <div class="bg-gradient-to-r from-purple-700 via-purple-800 to-indigo-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-7 h-7 sm:w-8 sm:h-8 rounded-xl bg-white text-purple-800 flex items-center justify-center font-black text-xs sm:text-sm flex-shrink-0 shadow-sm mt-0.5 sm:mt-0">2</div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Core Doctor Maximization</h4>
                                    <p class="text-[10px] sm:text-xs text-purple-100 mt-0.5 leading-tight">3 Doctors per Territory &bull; 1 Sweater Each</p>
                                </div>
                            </div>
                            <div class="self-start sm:self-auto pl-9 sm:pl-0">
                                <span class="text-[10px] sm:text-xs font-black bg-purple-950/80 text-purple-200 border border-purple-400/40 px-2.5 py-0.5 rounded-full inline-block whitespace-nowrap shadow-sm">3 Doctors Total</span>
                            </div>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
                            <!-- 2 Top Side-by-Side, 1 Centered in Middle Below -->
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4">
                                
                                <!-- Doc 1 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3.5 space-y-2.5">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">1</span> Doctor 1</span>
                                        <span id="c2_d1_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="space-y-2">
                                        <div>
                                            <label class="text-[10px] font-bold text-purple-950">Doctor 1 Name <span class="text-rose-500">*</span></label>
                                            <input type="text" id="c2_d1_name" oninput="onDataChanged()" placeholder="Enter Doctor 1 Name..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                        </div>
                                        <div>
                                            <div class="flex items-center justify-between"><label class="text-[10px] font-bold text-purple-950">Doctor 1 RPL ID <span class="text-rose-500">*</span></label><span id="c2_d1_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span></div>
                                            <input type="text" inputmode="numeric" maxlength="6" id="c2_d1_rpl" oninput="onRplInput(this, 'c2_d1_rpl_badge')" placeholder="6-digit RPL ID..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-purple-500 tracking-wider">
                                        </div>
                                    </div>
                                    <div class="flex gap-2.5 sm:gap-3 items-center pt-2 border-t border-purple-200/80">
                                        <div id="c2_d1_img_preview" onclick="zoomSlotImage('c2_d1_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                        <div class="flex-1 space-y-1.5 min-w-0">
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                <select id="c2_d1_sweater" onchange="onSweaterSelectChange('c2_d1', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-purple-500">
                                                    <option value="">-- Select Sweater --</option>
                                                    <option value="01 - Men's V-Neck (Grey)">01 - Men's V-Neck (Grey)</option>
                                                    <option value="02 - Men's V-Neck (Navy Blue)">02 - Men's V-Neck (Navy Blue)</option>
                                                    <option value="03 - Men's V-Neck (Cream Check)">03 - Men's V-Neck (Cream Check)</option>
                                                    <option value="04 - Women's Short Cardigan (Check)">04 - Women's Short Cardigan</option>
                                                    <option value="05 - Women's Semi Long Cardigan (Black)">05 - Women's Semi Long</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Size</label>
                                                <select id="c2_d1_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500"><option value="">-- Size --</option></select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Doc 2 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3.5 space-y-2.5">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">2</span> Doctor 2</span>
                                        <span id="c2_d2_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="space-y-2">
                                        <div>
                                            <label class="text-[10px] font-bold text-purple-950">Doctor 2 Name <span class="text-rose-500">*</span></label>
                                            <input type="text" id="c2_d2_name" oninput="onDataChanged()" placeholder="Enter Doctor 2 Name..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                        </div>
                                        <div>
                                            <div class="flex items-center justify-between"><label class="text-[10px] font-bold text-purple-950">Doctor 2 RPL ID <span class="text-rose-500">*</span></label><span id="c2_d2_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span></div>
                                            <input type="text" inputmode="numeric" maxlength="6" id="c2_d2_rpl" oninput="onRplInput(this, 'c2_d2_rpl_badge')" placeholder="6-digit RPL ID..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-purple-500 tracking-wider">
                                        </div>
                                    </div>
                                    <div class="flex gap-2.5 sm:gap-3 items-center pt-2 border-t border-purple-200/80">
                                        <div id="c2_d2_img_preview" onclick="zoomSlotImage('c2_d2_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                        <div class="flex-1 space-y-1.5 min-w-0">
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                <select id="c2_d2_sweater" onchange="onSweaterSelectChange('c2_d2', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-purple-500">
                                                    <option value="">-- Select Sweater --</option>
                                                    <option value="01 - Men's V-Neck (Grey)">01 - Men's V-Neck (Grey)</option>
                                                    <option value="02 - Men's V-Neck (Navy Blue)">02 - Men's V-Neck (Navy Blue)</option>
                                                    <option value="03 - Men's V-Neck (Cream Check)">03 - Men's V-Neck (Cream Check)</option>
                                                    <option value="04 - Women's Short Cardigan (Check)">04 - Women's Short Cardigan</option>
                                                    <option value="05 - Women's Semi Long Cardigan (Black)">05 - Women's Semi Long</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Size</label>
                                                <select id="c2_d2_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500"><option value="">-- Size --</option></select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Doc 3 (Middle Centered Below) -->
                                <div class="md:col-span-2 flex justify-center">
                                    <div class="w-full max-w-xl bg-purple-50/50 border border-purple-200 rounded-2xl p-3.5 space-y-2.5 shadow-sm">
                                        <div class="flex items-center justify-between">
                                            <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">3</span> Doctor 3</span>
                                            <span id="c2_d3_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                        </div>
                                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="text-[10px] font-bold text-purple-950">Doctor 3 Name <span class="text-rose-500">*</span></label>
                                                <input type="text" id="c2_d3_name" oninput="onDataChanged()" placeholder="Enter Doctor 3 Name..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                            </div>
                                            <div>
                                                <div class="flex items-center justify-between"><label class="text-[10px] font-bold text-purple-950">Doctor 3 RPL ID <span class="text-rose-500">*</span></label><span id="c2_d3_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span></div>
                                                <input type="text" inputmode="numeric" maxlength="6" id="c2_d3_rpl" oninput="onRplInput(this, 'c2_d3_rpl_badge')" placeholder="6-digit RPL ID..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-purple-500 tracking-wider">
                                            </div>
                                        </div>
                                        <div class="flex gap-2.5 sm:gap-3 items-center pt-2 border-t border-purple-200/80">
                                            <div id="c2_d3_img_preview" onclick="zoomSlotImage('c2_d3_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                            <div class="flex-1 space-y-1.5 min-w-0">
                                                <div>
                                                    <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                    <select id="c2_d3_sweater" onchange="onSweaterSelectChange('c2_d3', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-purple-500">
                                                        <option value="">-- Select Sweater --</option>
                                                        <option value="01 - Men's V-Neck (Grey)">01 - Men's V-Neck (Grey)</option>
                                                        <option value="02 - Men's V-Neck (Navy Blue)">02 - Men's V-Neck (Navy Blue)</option>
                                                        <option value="03 - Men's V-Neck (Cream Check)">03 - Men's V-Neck (Cream Check)</option>
                                                        <option value="04 - Women's Short Cardigan (Check)">04 - Women's Short Cardigan</option>
                                                        <option value="05 - Women's Semi Long Cardigan (Black)">05 - Women's Semi Long</option>
                                                    </select>
                                                </div>
                                                <div>
                                                    <label class="text-[10px] font-bold text-slate-500">Size</label>
                                                    <select id="c2_d3_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500"><option value="">-- Size --</option></select>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                            </div>
                        </div>
                    </div>

                    <!-- FOOTER CONTROLS & SAVE -->
                    <div class="bg-white rounded-3xl p-4 sm:p-5 border border-slate-200 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-3">
                        <div class="flex items-center gap-2 w-full sm:w-auto justify-between sm:justify-start">
                            <button onclick="navigateTerritory(-1)" class="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition flex items-center gap-1.5">
                                <i class="fa-solid fa-arrow-left"></i>
                                <span>Previous</span>
                            </button>
                            <span id="territory-step-indicator" class="text-xs font-bold text-slate-500">1 of 12</span>
                            <button onclick="navigateTerritory(1)" class="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition flex items-center gap-1.5">
                                <span>Next</span>
                                <i class="fa-solid fa-arrow-right"></i>
                            </button>
                        </div>

                        <div class="flex items-center gap-2.5 w-full sm:w-auto">
                            <button onclick="saveCurrentTerritoryClick()" class="flex-1 sm:flex-none px-6 py-2.5 bg-orange-500 hover:bg-orange-600 text-white rounded-xl text-xs font-black shadow-md hover:shadow-lg transition flex items-center justify-center gap-2 active:scale-95">
                                <i class="fa-solid fa-floppy-disk"></i>
                                <span>Save</span>
                            </button>
                        </div>
                    </div>

                </div>
            </div>
        </main>

    </div>

    <!-- CATALOGUE & MEASUREMENT MODAL -->
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

                <!-- Detailed Measurement Specs -->
                <div class="space-y-4 pt-2 border-t border-slate-200">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h4 class="text-sm sm:text-base font-black text-slate-900 flex items-center gap-2">
                                <i class="fa-solid fa-ruler-combined text-orange-500"></i>
                                <span>Sweater Size Measurement Chart (Inches)</span>
                            </h4>
                            <p class="text-[11px] text-slate-500">Standard apparel measurement specifications by Lubnan Trade Consortium Ltd. (Richman / Lubnan)</p>
                        </div>
                        <span class="text-[10px] font-bold bg-orange-100 text-orange-800 border border-orange-200 px-2.5 py-1 rounded-full self-start sm:self-auto">All Measurements in Inches (")</span>
                    </div>

                    <!-- Men's Sleeveless V-Neck -->
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
                                    <tr class="hover:bg-slate-50"><td class="p-2.5 text-left font-black text-orange-600">S</td><td class="p-2.5 font-bold text-slate-800">38"</td><td class="p-2.5">26"</td><td class="p-2.5">15"</td><td class="p-2.5 text-left text-slate-600">Slim / Lean Build</td></tr>
                                    <tr class="hover:bg-slate-50"><td class="p-2.5 text-left font-black text-orange-600">M</td><td class="p-2.5 font-bold text-slate-800">40"</td><td class="p-2.5">27"</td><td class="p-2.5">16"</td><td class="p-2.5 text-left text-slate-600">Medium Build (Standard)</td></tr>
                                    <tr class="hover:bg-slate-50"><td class="p-2.5 text-left font-black text-orange-600">L</td><td class="p-2.5 font-bold text-slate-800">42"</td><td class="p-2.5">28"</td><td class="p-2.5">17"</td><td class="p-2.5 text-left text-slate-600">Standard Adult Fit</td></tr>
                                    <tr class="hover:bg-slate-50"><td class="p-2.5 text-left font-black text-orange-600">XL</td><td class="p-2.5 font-bold text-slate-800">44"</td><td class="p-2.5">29"</td><td class="p-2.5">18"</td><td class="p-2.5 text-left text-slate-600">Plus / Comfort Fit</td></tr>
                                    <tr class="hover:bg-slate-50"><td class="p-2.5 text-left font-black text-orange-600">XXL</td><td class="p-2.5 font-bold text-slate-800">46"</td><td class="p-2.5">30"</td><td class="p-2.5">19"</td><td class="p-2.5 text-left text-slate-600">Extra Comfort / Loose Fit</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- Women's Cardigans Grid -->
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <!-- Short Cardigan -->
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
                                        <tr><th class="p-2 text-left">Size</th><th class="p-2">Chest</th><th class="p-2">Length</th><th class="p-2">Sleeve</th><th class="p-2">Shoulder</th></tr>
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

                        <!-- Semi Long Cardigan -->
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
                                        <tr><th class="p-2 text-left">Size</th><th class="p-2">Chest</th><th class="p-2">Length</th><th class="p-2">Sleeve</th><th class="p-2">Shoulder</th></tr>
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

    <!-- LIGHTBOX MODAL -->
    <div id="image-lightbox" class="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md hidden flex items-center justify-center p-4" onclick="closeImageLightbox()">
        <div class="bg-white rounded-3xl max-w-lg w-full overflow-hidden shadow-2xl border border-slate-800" onclick="event.stopPropagation()">
            <div class="p-3 bg-slate-900 text-white flex items-center justify-between">
                <span id="lightbox-title" class="text-xs font-bold">Sweater Preview</span>
                <button onclick="closeImageLightbox()" class="w-7 h-7 rounded-full bg-slate-800 text-slate-300 hover:text-white flex items-center justify-center"><i class="fa-solid fa-xmark text-sm"></i></button>
            </div>
            <div class="p-4 bg-slate-100 flex items-center justify-center">
                <img id="lightbox-img" src="" alt="Sweater" class="max-h-[60vh] object-contain rounded-2xl shadow-md border border-slate-200">
            </div>
            <div class="p-4 bg-white space-y-1">
                <div class="flex items-center justify-between">
                    <h4 id="lightbox-desc" class="font-black text-sm text-slate-900">Description</h4>
                    <span id="lightbox-gender" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-orange-100 text-orange-800">Gender</span>
                </div>
                <p id="lightbox-color" class="text-xs text-slate-500">Color Details</p>
            </div>
        </div>
    </div>

    <!-- ADMIN MODAL -->
    <div id="admin-modal" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-3 sm:p-6" onclick="closeAdminModal()">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-6xl w-full max-h-[92vh] flex flex-col overflow-hidden" onclick="event.stopPropagation()">
            <div class="px-5 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-900 text-white">
                <div class="flex items-center gap-2.5">
                    <div class="w-8 h-8 rounded-xl bg-orange-500 text-white flex items-center justify-center font-bold text-sm shadow-sm"><i class="fa-solid fa-shield-halved"></i></div>
                    <div>
                        <h3 class="font-bold text-xs sm:text-sm">Admin Control & Live Aggregation Portal</h3>
                        <p class="text-[10px] text-slate-400">National Campaign Dashboard &bull; 1,856 Total Territories</p>
                    </div>
                </div>
                <div class="flex items-center gap-2">
                    <button onclick="pullCloudData(true)" class="px-3 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold rounded-xl flex items-center gap-1.5 transition shadow-sm"><i class="fa-solid fa-rotate"></i><span>Pull Cloud Data</span></button>
                    <button onclick="closeAdminModal()" class="w-8 h-8 rounded-full bg-slate-800 text-slate-300 hover:text-white flex items-center justify-center transition"><i class="fa-solid fa-xmark text-sm"></i></button>
                </div>
            </div>

            <div class="p-4 sm:p-6 overflow-y-auto custom-scrollbar space-y-6 bg-slate-50">
                <!-- Master KPIs -->
                <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4">
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                        <span class="text-[10px] font-bold uppercase text-slate-400 tracking-wider">Total Territories</span>
                        <div id="admin-kpi-total" class="text-xl sm:text-2xl font-black text-slate-900">1,856</div>
                    </div>
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                        <span class="text-[10px] font-bold uppercase text-emerald-600 tracking-wider">Completed</span>
                        <div id="admin-kpi-complete" class="text-xl sm:text-2xl font-black text-emerald-600">0</div>
                    </div>
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                        <span class="text-[10px] font-bold uppercase text-amber-600 tracking-wider">In Progress</span>
                        <div id="admin-kpi-progress" class="text-xl sm:text-2xl font-black text-amber-600">0</div>
                    </div>
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                        <span class="text-[10px] font-bold uppercase text-orange-600 tracking-wider">Total Sweaters</span>
                        <div id="admin-kpi-sweaters" class="text-xl sm:text-2xl font-black text-orange-600">0</div>
                    </div>
                </div>

                <!-- Production Matrix -->
                <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
                    <div class="flex items-center justify-between">
                        <h4 class="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-900 flex items-center gap-2">
                            <i class="fa-solid fa-table-cells text-orange-500"></i>
                            <span>Sweater Production Breakdown Matrix</span>
                        </h4>
                    </div>
                    <div class="overflow-x-auto">
                        <table class="w-full text-xs text-center border-collapse">
                            <thead>
                                <tr class="bg-slate-900 text-white font-bold">
                                    <th class="p-2.5 text-left">Sweater Design</th>
                                    <th class="p-2.5">XS</th>
                                    <th class="p-2.5">S</th>
                                    <th class="p-2.5">M</th>
                                    <th class="p-2.5">L</th>
                                    <th class="p-2.5">XL</th>
                                    <th class="p-2.5">XXL</th>
                                    <th class="p-2.5 text-orange-400">Total</th>
                                </tr>
                            </thead>
                            <tbody id="admin-production-tbody" class="divide-y divide-slate-100">
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Zone Progress -->
                <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
                    <h4 class="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-900 flex items-center gap-2">
                        <i class="fa-solid fa-chart-pie text-teal-600"></i>
                        <span>Zone Requisition Progress (35 Zones)</span>
                    </h4>
                    <div id="admin-zone-progress-list" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    </div>
                </div>

                <!-- All 155 Regions Table -->
                <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <h4 class="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-900 flex items-center gap-2">
                            <i class="fa-solid fa-map-location-dot text-indigo-600"></i>
                            <span>All Regions Status (155 Regions)</span>
                        </h4>
                        <input type="text" id="admin-region-search" oninput="renderAdminRegionsTable(this.value)" placeholder="Search Region or Head..." class="bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-800 focus:outline-none focus:border-indigo-500 w-full sm:w-64">
                    </div>
                    <div class="overflow-x-auto max-h-80 custom-scrollbar border border-slate-200 rounded-xl">
                        <table class="w-full text-xs text-left">
                            <thead class="bg-slate-900 text-white font-bold sticky top-0">
                                <tr>
                                    <th class="p-2.5">Region</th>
                                    <th class="p-2.5">Regional Head</th>
                                    <th class="p-2.5">Zone</th>
                                    <th class="p-2.5 text-center">Territories</th>
                                    <th class="p-2.5 text-center">Status</th>
                                    <th class="p-2.5 text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody id="admin-regions-tbody" class="divide-y divide-slate-100">
                            </tbody>
                        </table>
                    </div>
                </div>

                <!-- Master Actions -->
                <div class="bg-slate-900 text-white rounded-2xl p-4 sm:p-5 space-y-4">
                    <h4 class="text-xs font-black uppercase tracking-wider text-orange-400">Master Actions & Live Cloud Operations</h4>
                    <div class="flex flex-wrap items-center gap-3">
                        <button onclick="exportMasterExcelFromAdmin()" class="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black flex items-center gap-2 shadow-sm transition active:scale-95">
                            <i class="fa-solid fa-file-excel text-base"></i>
                            <span>Export Live Master Excel</span>
                        </button>
                        <button onclick="pushAllStoredDataToGoogleSheet()" class="px-4 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-black flex items-center gap-2 shadow-sm transition active:scale-95">
                            <i class="fa-solid fa-cloud-arrow-up text-base"></i>
                            <span>Push All to Google Sheet</span>
                        </button>
                        <button onclick="toggleCloudSettings()" class="px-4 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 rounded-xl text-xs font-bold flex items-center gap-2 transition">
                            <i class="fa-solid fa-gear"></i>
                            <span>Cloud Settings</span>
                        </button>
                        <button onclick="deleteAllCampaignData()" class="px-4 py-2.5 bg-rose-950/80 hover:bg-rose-900 text-rose-300 border border-rose-800 rounded-xl text-xs font-bold flex items-center gap-2 transition ml-auto">
                            <i class="fa-solid fa-trash-can"></i>
                            <span>Delete All Data</span>
                        </button>
                    </div>

                    <!-- Cloud Settings Drawer -->
                    <div id="cloud-settings-drawer" class="hidden pt-4 border-t border-slate-800 space-y-3">
                        <label class="block text-xs font-bold text-slate-300">Google Apps Script Web App URL</label>
                        <div class="flex gap-2">
                            <input type="text" id="custom-cloud-url-input" class="flex-1 bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 font-mono focus:outline-none focus:border-indigo-500" placeholder="https://script.google.com/macros/s/.../exec">
                            <button onclick="saveCloudUrlSetting()" class="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold rounded-xl transition">Save URL</button>
                            <button onclick="testCloudConnection()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-bold rounded-xl transition">Test Connection</button>
                        </div>
                        <div id="cloud-test-result" class="hidden text-xs font-bold p-2.5 rounded-xl border border-slate-700"></div>
                    </div>
                </div>

            </div>
        </div>
    </div>

    <!-- TOAST NOTIFICATION -->
    <div id="toast" class="fixed bottom-5 right-5 z-50 transform translate-y-20 opacity-0 transition-all duration-300 pointer-events-none">
        <div id="toast-body" class="bg-slate-900 text-white px-4 py-3 rounded-2xl shadow-2xl border border-slate-700 flex items-center gap-2.5 text-xs font-bold">
            <span id="toast-msg">Notification</span>
        </div>
    </div>

    <!-- JAVASCRIPT LOGIC -->
    <script>
        const REGION_MAP = ###REGION_MAP###;
        const ALL_TERRITORIES = ###ALL_TERRITORIES###;
        const ZONES = ###ZONES###;
        const DEFAULT_CLOUD_URL = "###DEFAULT_CLOUD_URL###";

        const SWEATER_DETAILS = {
            "01": { code: "01", name: "Men's Sleeveless V-Neck Sweater", color: "Solid Ash / Grey Textured", gender: "Men's", sizes: ["S", "M", "L", "XL", "XXL"] },
            "02": { code: "02", name: "Men's Sleeveless V-Neck Sweater", color: "Solid Navy Blue Textured", gender: "Men's", sizes: ["S", "M", "L", "XL", "XXL"] },
            "03": { code: "03", name: "Men's Sleeveless V-Neck Sweater", color: "Off-White / Cream Checkered", gender: "Men's", sizes: ["S", "M", "L", "XL", "XXL"] },
            "04": { code: "04", name: "Women's Short Cardigan", color: "White & Navy Grid Check", gender: "Women's", sizes: ["XS", "S", "M", "L", "XL"] },
            "05": { code: "05", name: "Women's Semi Long Cardigan", color: "Solid Black with Border Trim", gender: "Women's", sizes: ["S", "M", "L", "XL", "XXL"] }
        };

        let store = JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{}');
        let isGlobalAccessOpen = JSON.parse(localStorage.getItem('EXIUM_GLOBAL_ACCESS') || 'true');
        let cloudApiUrl = localStorage.getItem('EXIUM_CLOUD_URL') || DEFAULT_CLOUD_URL;

        let currentRegionCode = null;
        let activeTerritoryIndex = 0;
        let isAdminLoggedIn = false;
        let autoSyncTimeout = null;

        window.addEventListener('DOMContentLoaded', () => {
            populateZoneDropdown();
            
            const savedSession = JSON.parse(localStorage.getItem('EXIUM_ACTIVE_SESSION') || 'null');
            if (savedSession && savedSession.region_code && REGION_MAP[savedSession.region_code]) {
                unlockRegion(savedSession.region_code, true);
                if (typeof savedSession.territory_idx === 'number') {
                    selectTerritoryTab(savedSession.territory_idx, false);
                }
            }
        });

        // -------------------------------------------------------------
        // SELECTION & LOGIN
        // -------------------------------------------------------------
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
            const passCont = document.getElementById('password-container');
            const btn = document.getElementById('btn-unlock-region');

            regSel.innerHTML = '<option value="">-- Select Your Region --</option>';
            passCont.classList.add('hidden');
            btn.disabled = true;

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
                    opt.textContent = `${r.region_name} (${r.regional_head}) - SAP: ${r.sap_region_code}`;
                    regSel.appendChild(opt);
                }
            }
        }

        function onRegionChanged() {
            const regCode = document.getElementById('select-region').value;
            const passCont = document.getElementById('password-container');
            const btn = document.getElementById('btn-unlock-region');

            if (regCode) {
                passCont.classList.remove('hidden');
                btn.disabled = false;
                document.getElementById('region-password').focus();
            } else {
                passCont.classList.add('hidden');
                btn.disabled = true;
            }
        }

        function handlePasswordKey(e) {
            if (e.key === 'Enter') {
                unlockRegion();
            }
        }

        function unlockRegion(directCode = null, isRestoringSession = false) {
            const regCode = directCode || document.getElementById('select-region').value;
            if (!regCode || !REGION_MAP[regCode]) {
                alert("Please select a valid Region!");
                return;
            }

            if (!isRestoringSession) {
                const pass = document.getElementById('region-password').value.trim();
                if (pass !== '1234' && pass !== 'admin') {
                    alert("Invalid PIN / Password! Please enter the correct PIN (Default: 1234)");
                    return;
                }
            }

            currentRegionCode = regCode;
            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify({
                region_code: regCode,
                territory_idx: 0
            }));

            const r = REGION_MAP[regCode];
            document.getElementById('selection-view').classList.add('hidden');
            document.getElementById('workspace-view').classList.remove('hidden');

            document.getElementById('banner-region').textContent = `SAP: ${r.sap_region_code}`;
            document.getElementById('banner-rh').textContent = `Region: ${r.region_name} (${r.regional_head})`;
            document.getElementById('sidebar-count').textContent = `${r.territories.length} Total`;

            activeTerritoryIndex = 0;
            renderTerritoryTabs();
            selectTerritoryTab(0, true);
        }

        function exitRegionWorkspace() {
            try {
                onDataChanged();
            } catch (e) {
                console.warn(e);
            }
            localStorage.removeItem('EXIUM_ACTIVE_SESSION');
            currentRegionCode = null;
            document.getElementById('workspace-view').classList.add('hidden');
            document.getElementById('selection-view').classList.remove('hidden');
            const passInput = document.getElementById('region-password');
            if (passInput) passInput.value = '';
        }

        // -------------------------------------------------------------
        // TERRITORY NAVIGATION & TAB RENDERING (3 DISTINCT COLORS)
        // -------------------------------------------------------------
        function renderTerritoryTabs() {
            if (!currentRegionCode || !REGION_MAP[currentRegionCode]) return;
            const r = REGION_MAP[currentRegionCode];
            const deskList = document.getElementById('desktop-territory-list');
            const mobSelect = document.getElementById('mobile-territory-select');

            deskList.innerHTML = '';
            mobSelect.innerHTML = '';

            let completedCount = 0;

            r.territories.forEach((t, idx) => {
                const d = store[String(t.sap_territory_code)] || {};
                const status = getTerritoryStatus(d);
                if (status === 'Complete') completedCount++;

                const mobOpt = document.createElement('option');
                mobOpt.value = idx;
                mobOpt.textContent = `${t.territory_name} (${status})`;
                if (idx === activeTerritoryIndex) mobOpt.selected = true;
                mobSelect.appendChild(mobOpt);

                const isActive = (idx === activeTerritoryIndex);

                let btnClasses = '';
                let badgeClasses = '';
                let badgeText = status;

                if (status === 'Complete') {
                    badgeText = '✓ Complete';
                    if (isActive) {
                        btnClasses = 'bg-emerald-600 text-white border-emerald-600 shadow-md ring-2 ring-emerald-500/30';
                        badgeClasses = 'bg-white text-emerald-900 font-black shadow-sm';
                    } else {
                        btnClasses = 'bg-emerald-50/70 hover:bg-emerald-100/70 text-emerald-950 border-emerald-200/90';
                        badgeClasses = 'bg-emerald-100 text-emerald-800 border border-emerald-300 font-bold';
                    }
                } else if (status === 'In Progress') {
                    badgeText = '⏳ In Progress';
                    if (isActive) {
                        btnClasses = 'bg-amber-500 text-white border-amber-500 shadow-md ring-2 ring-amber-500/30';
                        badgeClasses = 'bg-white text-amber-900 font-black shadow-sm';
                    } else {
                        btnClasses = 'bg-amber-50/70 hover:bg-amber-100/70 text-amber-950 border-amber-200/90';
                        badgeClasses = 'bg-amber-100 text-amber-800 border border-amber-300 font-bold';
                    }
                } else {
                    // Not Started
                    badgeText = '○ Not Started';
                    if (isActive) {
                        btnClasses = 'bg-slate-900 text-white border-slate-900 shadow-md ring-2 ring-slate-900/30';
                        badgeClasses = 'bg-slate-700 text-slate-100 font-bold';
                    } else {
                        btnClasses = 'bg-white hover:bg-slate-50 text-slate-700 border-slate-200';
                        badgeClasses = 'bg-slate-100 text-slate-500 border border-slate-200 font-medium';
                    }
                }

                const btn = document.createElement('button');
                btn.type = 'button';
                btn.onclick = () => selectTerritoryTab(idx);
                btn.id = `terr-tab-btn-${idx}`;
                btn.className = `w-full text-left p-2.5 sm:p-3 rounded-2xl text-xs font-bold transition flex items-center justify-between border ${btnClasses}`;

                btn.innerHTML = `
                    <div class="truncate pr-1 min-w-0">
                        <div class="truncate font-black text-xs ${isActive ? 'text-white' : 'text-slate-900'}">${t.territory_name}</div>
                        <div class="text-[10px] ${isActive ? 'text-white/80' : 'text-slate-400'} font-mono">SAP: ${t.sap_territory_code}</div>
                    </div>
                    <span class="text-[9px] px-2 py-0.5 rounded-full flex-shrink-0 whitespace-nowrap ${badgeClasses}">${badgeText}</span>
                `;
                deskList.appendChild(btn);
            });

            document.getElementById('region-progress-badge').textContent = `${completedCount}/${r.territories.length} Done`;
        }

        function selectTerritoryTab(idx, shouldScroll = true) {
            activeTerritoryIndex = idx;
            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[idx];
            const terrCode = String(t.sap_territory_code);
            const d = store[terrCode] || {};

            renderTerritoryTabs();

            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify({
                region_code: currentRegionCode,
                territory_idx: idx
            }));

            document.getElementById('mobile-territory-select').value = idx;
            document.getElementById('current-territory-title').textContent = t.territory_name;
            document.getElementById('current-territory-code').textContent = `SAP Code: ${terrCode}`;
            document.getElementById('territory-step-indicator').textContent = `${idx + 1} of ${r.territories.length}`;

            const status = getTerritoryStatus(d);
            const statusBadge = document.getElementById('current-territory-status');
            statusBadge.textContent = status;
            statusBadge.className = `text-[10px] sm:text-xs font-black px-3 py-1 rounded-full ${
                status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-black' :
                'bg-white/10 text-slate-200 border border-white/20'
            }`;

            // Load Campaign 1
            const c1DocInput = document.getElementById('c1_doc_name');
            c1DocInput.value = d.c1_doc_name || '';
            const c1DocRpl = document.getElementById('c1_doc_rpl');
            c1DocRpl.value = d.c1_doc_rpl || '';
            updateRplBadgeState(c1DocRpl, 'c1_doc_rpl_badge');

            ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                const sw = d[`c1_${m}_sweater`] || '';
                const sz = d[`c1_${m}_size`] || '';
                const swEl = document.getElementById(`c1_${m}_sweater`);
                const szEl = document.getElementById(`c1_${m}_size`);
                if (swEl) swEl.value = sw;
                updateSizeOptionsForSelect(`c1_${m}_sweater`, `c1_${m}_size`, sz);
                updateSlotImagePreview(`c1_${m}`);
            });

            if (d.c1_m4_sweater || d.c1_m4_size) {
                showC1Sweater4(false);
            } else {
                hideC1Sweater4ViewOnly();
            }

            // Load Campaign 2 (3 Doctors)
            ['d1', 'd2', 'd3'].forEach(d_item => {
                const nameEl = document.getElementById(`c2_${d_item}_name`);
                const rplEl = document.getElementById(`c2_${d_item}_rpl`);
                const swEl = document.getElementById(`c2_${d_item}_sweater`);
                if (nameEl) nameEl.value = d[`c2_${d_item}_name`] || '';
                if (rplEl) {
                    rplEl.value = d[`c2_${d_item}_rpl`] || '';
                    updateRplBadgeState(rplEl, `c2_${d_item}_rpl_badge`);
                }
                if (swEl) swEl.value = d[`c2_${d_item}_sweater`] || '';
                updateSizeOptionsForSelect(`c2_${d_item}_sweater`, `c2_${d_item}_size`, d[`c2_${d_item}_size`] || '');
                updateSlotImagePreview(`c2_${d_item}`);
            });

            updateTerritorySlotCheckBadges(d);

            if (shouldScroll) {
                const bannerEl = document.getElementById('active-territory-banner-card');
                if (bannerEl && window.innerWidth < 1024) {
                    bannerEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            }
        }

        function navigateTerritory(delta) {
            const r = REGION_MAP[currentRegionCode];
            const nextIdx = activeTerritoryIndex + delta;
            if (nextIdx >= 0 && nextIdx < r.territories.length) {
                selectTerritoryTab(nextIdx, true);
            }
        }

        // -------------------------------------------------------------
        // FORM INTERACTIONS & STATUS LOGIC
        // -------------------------------------------------------------
        function onRplInput(input, badgeId) {
            input.value = input.value.replace(/[^0-9]/g, '');
            updateRplBadgeState(input, badgeId);
            onDataChanged();
        }

        function updateRplBadgeState(input, badgeId) {
            const badge = document.getElementById(badgeId);
            if (!badge) return;
            const len = input.value.length;
            if (len === 6) {
                badge.textContent = '✓ 6 digits';
                badge.className = 'text-[9px] font-bold text-emerald-600 bg-emerald-50 px-1.5 py-0.5 rounded';
            } else if (len > 0) {
                badge.textContent = `${len}/6 digits`;
                badge.className = 'text-[9px] font-bold text-amber-600 bg-amber-50 px-1.5 py-0.5 rounded';
            } else {
                badge.textContent = '6 digits';
                badge.className = 'text-[9px] font-bold text-slate-400';
            }
        }

        function onSweaterSelectChange(slotPrefix, sweaterVal) {
            updateSizeOptionsForSelect(`${slotPrefix}_sweater`, `${slotPrefix}_size`, '');
            updateSlotImagePreview(slotPrefix);
            onDataChanged();
        }

        function updateSizeOptionsForSelect(swId, szId, currentVal) {
            const swEl = document.getElementById(swId);
            const szEl = document.getElementById(szId);
            if (!swEl || !szEl) return;

            const code = (swEl.value || '').substring(0, 2);
            szEl.innerHTML = '<option value="">-- Size --</option>';

            if (SWEATER_DETAILS[code]) {
                SWEATER_DETAILS[code].sizes.forEach(s => {
                    const opt = document.createElement('option');
                    opt.value = s;
                    opt.textContent = s;
                    if (s === currentVal) opt.selected = true;
                    szEl.appendChild(opt);
                });
            }
        }

        function updateSlotImagePreview(slotPrefix) {
            const swEl = document.getElementById(`${slotPrefix}_sweater`);
            const prevEl = document.getElementById(`${slotPrefix}_img_preview`);
            if (!swEl || !prevEl) return;

            const code = (swEl.value || '').substring(0, 2);
            if (SWEATER_DETAILS[code]) {
                const imgB64 = code === '01' ? '###B64_01###' : code === '02' ? '###B64_02###' : code === '03' ? '###B64_03###' : code === '04' ? '###B64_04###' : '###B64_05###';
                prevEl.innerHTML = `<img src="${imgB64}" class="w-full h-full object-cover"><span class="absolute top-1 left-1 bg-slate-900/90 text-white text-[9px] font-black px-1.5 py-0.2 rounded">${code}</span>`;
            } else {
                prevEl.innerHTML = `<i class="fa-solid fa-shirt text-lg text-slate-300"></i>`;
            }
        }

        function showC1Sweater4(triggerChange = false) {
            const container = document.getElementById('c1_m4_container');
            const btn = document.getElementById('c1_add_m4_btn_container');
            if (container) container.classList.remove('hidden');
            if (btn) btn.classList.add('hidden');
            if (triggerChange) onDataChanged();
        }

        function hideC1Sweater4ViewOnly() {
            const container = document.getElementById('c1_m4_container');
            const btn = document.getElementById('c1_add_m4_btn_container');
            if (container) container.classList.add('hidden');
            if (btn) btn.classList.remove('hidden');
        }

        function hideAndClearC1Sweater4() {
            const swSelect = document.getElementById('c1_m4_sweater');
            const szSelect = document.getElementById('c1_m4_size');
            if (swSelect) swSelect.value = '';
            if (szSelect) szSelect.innerHTML = '<option value="">-- Size --</option>';
            updateSlotImagePreview('c1_m4');
            hideC1Sweater4ViewOnly();
            onDataChanged();
        }

        function onDataChanged() {
            if (!currentRegionCode || !REGION_MAP[currentRegionCode]) return;
            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[activeTerritoryIndex];
            const terrCode = String(t.sap_territory_code);

            const isM4Visible = !document.getElementById('c1_m4_container')?.classList.contains('hidden');

            const terrData = {
                c1_doc_name: document.getElementById('c1_doc_name')?.value.trim() || '',
                c1_doc_rpl: document.getElementById('c1_doc_rpl')?.value.trim() || '',
                c1_m1_sweater: document.getElementById('c1_m1_sweater')?.value || '',
                c1_m1_size: document.getElementById('c1_m1_size')?.value || '',
                c1_m2_sweater: document.getElementById('c1_m2_sweater')?.value || '',
                c1_m2_size: document.getElementById('c1_m2_size')?.value || '',
                c1_m3_sweater: document.getElementById('c1_m3_sweater')?.value || '',
                c1_m3_size: document.getElementById('c1_m3_size')?.value || '',
                c1_m4_sweater: isM4Visible ? (document.getElementById('c1_m4_sweater')?.value || '') : '',
                c1_m4_size: isM4Visible ? (document.getElementById('c1_m4_size')?.value || '') : '',

                c2_d1_name: document.getElementById('c2_d1_name')?.value.trim() || '',
                c2_d1_rpl: document.getElementById('c2_d1_rpl')?.value.trim() || '',
                c2_d1_sweater: document.getElementById('c2_d1_sweater')?.value || '',
                c2_d1_size: document.getElementById('c2_d1_size')?.value || '',

                c2_d2_name: document.getElementById('c2_d2_name')?.value.trim() || '',
                c2_d2_rpl: document.getElementById('c2_d2_rpl')?.value.trim() || '',
                c2_d2_sweater: document.getElementById('c2_d2_sweater')?.value || '',
                c2_d2_size: document.getElementById('c2_d2_size')?.value || '',

                c2_d3_name: document.getElementById('c2_d3_name')?.value.trim() || '',
                c2_d3_rpl: document.getElementById('c2_d3_rpl')?.value.trim() || '',
                c2_d3_sweater: document.getElementById('c2_d3_sweater')?.value || '',
                c2_d3_size: document.getElementById('c2_d3_size')?.value || '',

                c2_d4_name: '',
                c2_d4_rpl: '',
                c2_d4_sweater: '',
                c2_d4_size: ''
            };

            store[terrCode] = terrData;
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));

            updateTerritorySlotCheckBadges(terrData);
            renderTerritoryTabs();

            const status = getTerritoryStatus(terrData);
            const statusBadge = document.getElementById('current-territory-status');
            if (statusBadge) {
                statusBadge.textContent = status;
                statusBadge.className = `text-[10px] sm:text-xs font-black px-3 py-1 rounded-full ${
                    status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                    status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-black' :
                    'bg-white/10 text-slate-200 border border-white/20'
                }`;
            }

            // Debounced Cloud Auto-Sync
            if (autoSyncTimeout) clearTimeout(autoSyncTimeout);
            autoSyncTimeout = setTimeout(() => {
                syncTerritoryToCloud(terrCode, terrData);
            }, 1200);
        }

        function getTerritoryStatus(d) {
            if (!d) return 'Not Started';

            // Campaign 1 Check
            const c1Mandatory = Boolean(d.c1_doc_name && d.c1_doc_rpl && String(d.c1_doc_rpl).length === 6 &&
                                        d.c1_m1_sweater && d.c1_m1_size &&
                                        d.c1_m2_sweater && d.c1_m2_size &&
                                        d.c1_m3_sweater && d.c1_m3_size);
            const c1M4HasAny = Boolean(d.c1_m4_sweater || d.c1_m4_size);
            const c1M4Ok = !c1M4HasAny || Boolean(d.c1_m4_sweater && d.c1_m4_size);
            const c1Ok = c1Mandatory && c1M4Ok;

            // Campaign 2 Check (3 Doctors)
            const c2Doc1Ok = Boolean(d.c2_d1_name && d.c2_d1_rpl && String(d.c2_d1_rpl).length === 6 && d.c2_d1_sweater && d.c2_d1_size);
            const c2Doc2Ok = Boolean(d.c2_d2_name && d.c2_d2_rpl && String(d.c2_d2_rpl).length === 6 && d.c2_d2_sweater && d.c2_d2_size);
            const c2Doc3Ok = Boolean(d.c2_d3_name && d.c2_d3_rpl && String(d.c2_d3_rpl).length === 6 && d.c2_d3_sweater && d.c2_d3_size);
            const c2Ok = c2Doc1Ok && c2Doc2Ok && c2Doc3Ok;

            if (c1Ok && c2Ok) return 'Complete';

            // In Progress Check
            const hasAny = Boolean(d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater || d.c1_m2_sweater || d.c1_m3_sweater || d.c1_m4_sweater ||
                                   d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater ||
                                   d.c2_d2_name || d.c2_d2_rpl || d.c2_d2_sweater ||
                                   d.c2_d3_name || d.c2_d3_rpl || d.c2_d3_sweater);
            
            return hasAny ? 'In Progress' : 'Not Started';
        }

        function updateTerritorySlotCheckBadges(d) {
            if (!d) return;

            // C1 Slots
            ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                const sw = d[`c1_${m}_sweater`];
                const sz = d[`c1_${m}_size`];
                const badge = document.getElementById(`c1_${m}_check_badge`);
                if (badge) {
                    if (sw && sz) {
                        badge.innerHTML = '<span class="text-emerald-600 text-[10px] font-bold"><i class="fa-solid fa-circle-check"></i> Complete</span>';
                    } else if (sw || sz) {
                        badge.innerHTML = '<span class="text-amber-500 text-[10px] font-bold"><i class="fa-solid fa-clock"></i> Partial</span>';
                    } else {
                        badge.innerHTML = '<span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span>';
                    }
                }
            });

            // C2 Slots (3 Doctors)
            ['d1', 'd2', 'd3'].forEach(d_item => {
                const name = d[`c2_${d_item}_name`];
                const rpl = d[`c2_${d_item}_rpl`];
                const sw = d[`c2_${d_item}_sweater`];
                const sz = d[`c2_${d_item}_size`];
                const badge = document.getElementById(`c2_${d_item}_check_badge`);
                if (badge) {
                    if (name && rpl && String(rpl).length === 6 && sw && sz) {
                        badge.innerHTML = '<span class="text-emerald-600 text-[10px] font-bold"><i class="fa-solid fa-circle-check"></i> Complete</span>';
                    } else if (name || rpl || sw || sz) {
                        badge.innerHTML = '<span class="text-amber-500 text-[10px] font-bold"><i class="fa-solid fa-clock"></i> Partial</span>';
                    } else {
                        badge.innerHTML = '<span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span>';
                    }
                }
            });
        }

        function saveCurrentTerritoryClick() {
            onDataChanged();
            showToast("💾 Saved! Synced to local & cloud storage.");
        }

        function showToast(msg) {
            const toast = document.getElementById('toast');
            const toastMsg = document.getElementById('toast-msg');
            if (!toast || !toastMsg) return;
            toastMsg.textContent = msg;
            toast.classList.remove('translate-y-20', 'opacity-0');
            setTimeout(() => {
                toast.classList.add('translate-y-20', 'opacity-0');
            }, 3000);
        }

        // -------------------------------------------------------------
        // MODALS & LIGHTBOX
        // -------------------------------------------------------------
        function openImageLightbox(code) {
            const item = SWEATER_DETAILS[code];
            if (!item) return;
            const imgB64 = code === '01' ? '###B64_01###' : code === '02' ? '###B64_02###' : code === '03' ? '###B64_03###' : code === '04' ? '###B64_04###' : '###B64_05###';
            document.getElementById('lightbox-img').src = imgB64;
            document.getElementById('lightbox-title').textContent = `Design Code: ${item.code}`;
            document.getElementById('lightbox-desc').textContent = item.name;
            document.getElementById('lightbox-gender').textContent = item.gender;
            document.getElementById('lightbox-color').textContent = `Color: ${item.color} | Available Sizes: ${item.sizes.join(', ')}`;
            document.getElementById('image-lightbox').classList.remove('hidden');
        }

        function zoomSlotImage(slotSelectId) {
            const sel = document.getElementById(slotSelectId);
            if (!sel || !sel.value) return;
            const code = sel.value.substring(0, 2);
            if (SWEATER_DETAILS[code]) openImageLightbox(code);
        }

        function closeImageLightbox() {
            document.getElementById('image-lightbox').classList.add('hidden');
        }

        function openCatalogModal() {
            document.getElementById('catalog-modal').classList.remove('hidden');
        }

        function closeCatalogModal() {
            document.getElementById('catalog-modal').classList.add('hidden');
        }

        function togglePasswordVisibility(inputId) {
            const el = document.getElementById(inputId);
            if (!el) return;
            el.type = el.type === 'password' ? 'text' : 'password';
        }

        // -------------------------------------------------------------
        // ADMIN DASHBOARD & LIVE AGGREGATION
        // -------------------------------------------------------------
        function openAdminModal() {
            document.getElementById('admin-modal').classList.remove('hidden');
            renderAdminKpisAndSummaries();
            renderAdminZoneProgress();
            renderAdminProductionMatrix();
            renderAdminRegionsTable();
            pullCloudData(false);
        }

        function closeAdminModal() {
            document.getElementById('admin-modal').classList.add('hidden');
        }

        function renderAdminKpisAndSummaries() {
            let total = ALL_TERRITORIES.length;
            let comp = 0;
            let inProg = 0;
            let totalSweaters = 0;

            ALL_TERRITORIES.forEach(t => {
                const d = store[String(t['SAP Territory Code'])] || {};
                const st = getTerritoryStatus(d);
                if (st === 'Complete') comp++;
                else if (st === 'In Progress') inProg++;

                // Count sweaters in C1
                ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                    if (d[`c1_${m}_sweater`] && d[`c1_${m}_size`]) totalSweaters++;
                });
                // Count sweaters in C2 (3 Doctors)
                ['d1', 'd2', 'd3'].forEach(d_item => {
                    if (d[`c2_${d_item}_sweater`] && d[`c2_${d_item}_size`]) totalSweaters++;
                });
            });

            document.getElementById('admin-kpi-total').textContent = total.toLocaleString();
            document.getElementById('admin-kpi-complete').textContent = comp.toLocaleString();
            document.getElementById('admin-kpi-progress').textContent = inProg.toLocaleString();
            document.getElementById('admin-kpi-sweaters').textContent = totalSweaters.toLocaleString();
        }

        function renderAdminProductionMatrix() {
            const matrix = {
                '01': { name: "01 - Men's V-Neck (Grey)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                '02': { name: "02 - Men's V-Neck (Navy Blue)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                '03': { name: "03 - Men's V-Neck (Cream Check)", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                '04': { name: "04 - Women's Short Cardigan", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                '05': { name: "05 - Women's Semi Long Cardigan", XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }
            };

            for (let terrCode in store) {
                const d = store[terrCode];
                if (!d) continue;

                // C1
                ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                    const sw = d[`c1_${m}_sweater`];
                    const sz = d[`c1_${m}_size`];
                    if (sw && sz) {
                        const code = sw.substring(0, 2);
                        if (matrix[code] && matrix[code][sz] !== undefined) {
                            matrix[code][sz]++;
                            matrix[code].total++;
                        }
                    }
                });

                // C2 (3 Doctors)
                ['d1', 'd2', 'd3'].forEach(d_item => {
                    const sw = d[`c2_${d_item}_sweater`];
                    const sz = d[`c2_${d_item}_size`];
                    if (sw && sz) {
                        const code = sw.substring(0, 2);
                        if (matrix[code] && matrix[code][sz] !== undefined) {
                            matrix[code][sz]++;
                            matrix[code].total++;
                        }
                    }
                });
            }

            const tbody = document.getElementById('admin-production-tbody');
            tbody.innerHTML = '';

            let grandTotal = 0;
            let sizeTotals = { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0 };

            ['01', '02', '03', '04', '05'].forEach(code => {
                const c = matrix[code];
                grandTotal += c.total;
                for (let s in sizeTotals) sizeTotals[s] += c[s];

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="p-2.5 text-left font-bold text-slate-900">${c.name}</td>
                    <td class="p-2.5">${c.XS || '-'}</td>
                    <td class="p-2.5">${c.S || '-'}</td>
                    <td class="p-2.5">${c.M || '-'}</td>
                    <td class="p-2.5">${c.L || '-'}</td>
                    <td class="p-2.5">${c.XL || '-'}</td>
                    <td class="p-2.5">${c.XXL || '-'}</td>
                    <td class="p-2.5 font-black text-orange-600 bg-orange-50">${c.total}</td>
                `;
                tbody.appendChild(tr);
            });

            const trTotal = document.createElement('tr');
            trTotal.className = 'bg-slate-100 font-black text-slate-900 border-t-2 border-slate-300';
            trTotal.innerHTML = `
                <td class="p-2.5 text-left uppercase text-slate-700">Total Count</td>
                <td class="p-2.5">${sizeTotals.XS}</td>
                <td class="p-2.5">${sizeTotals.S}</td>
                <td class="p-2.5">${sizeTotals.M}</td>
                <td class="p-2.5">${sizeTotals.L}</td>
                <td class="p-2.5">${sizeTotals.XL}</td>
                <td class="p-2.5">${sizeTotals.XXL}</td>
                <td class="p-2.5 text-emerald-600 bg-emerald-100 text-sm">${grandTotal}</td>
            `;
            tbody.appendChild(trTotal);
        }

        function renderAdminZoneProgress() {
            const list = document.getElementById('admin-zone-progress-list');
            list.innerHTML = '';

            ZONES.forEach(zone => {
                let zTotal = 0;
                let zComp = 0;

                ALL_TERRITORIES.forEach(t => {
                    if (t.Zone === zone) {
                        zTotal++;
                        const d = store[String(t['SAP Territory Code'])] || {};
                        if (getTerritoryStatus(d) === 'Complete') zComp++;
                    }
                });

                const pct = zTotal > 0 ? Math.round((zComp / zTotal) * 100) : 0;

                const card = document.createElement('div');
                card.className = 'bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-1.5';
                card.innerHTML = `
                    <div class="flex items-center justify-between text-xs">
                        <span class="font-bold text-slate-900 truncate">${zone}</span>
                        <span class="font-black text-emerald-600">${zComp}/${zTotal}</span>
                    </div>
                    <div class="w-full bg-slate-200 h-1.5 rounded-full overflow-hidden">
                        <div class="bg-emerald-500 h-full rounded-full transition-all duration-300" style="width: ${pct}%"></div>
                    </div>
                `;
                list.appendChild(card);
            });
        }

        function renderAdminRegionsTable(query = '') {
            const tbody = document.getElementById('admin-regions-tbody');
            tbody.innerHTML = '';

            const q = query.toLowerCase();

            for (let code in REGION_MAP) {
                const r = REGION_MAP[code];
                if (q && !r.region_name.toLowerCase().includes(q) && !r.regional_head.toLowerCase().includes(q) && !r.zone.toLowerCase().includes(q)) {
                    continue;
                }

                let total = r.territories.length;
                let comp = 0;

                r.territories.forEach(t => {
                    const d = store[String(t.sap_territory_code)] || {};
                    if (getTerritoryStatus(d) === 'Complete') comp++;
                });

                const st = comp === total ? 'Complete' : (comp > 0 ? 'In Progress' : 'Not Started');

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="p-2.5 font-bold text-slate-900">${r.region_name}</td>
                    <td class="p-2.5 text-slate-600">${r.regional_head}</td>
                    <td class="p-2.5 text-slate-500">${r.zone}</td>
                    <td class="p-2.5 text-center font-black">${comp}/${total}</td>
                    <td class="p-2.5 text-center">
                        <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${
                            st === 'Complete' ? 'bg-emerald-100 text-emerald-800' :
                            st === 'In Progress' ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-500'
                        }">${st}</span>
                    </td>
                    <td class="p-2.5 text-center">
                        <button onclick="deleteSingleRegionData('${code}')" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-lg text-[10px] font-bold transition">Clear</button>
                    </td>
                `;
                tbody.appendChild(tr);
            }
        }

        async function deleteSingleRegionData(regCode) {
            const r = REGION_MAP[regCode];
            if (!r) return;
            if (!confirm(`Are you sure you want to clear all data for ${r.region_name} (${regCode})?`)) return;

            r.territories.forEach(t => {
                delete store[String(t.sap_territory_code)];
            });
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));

            // Sync region deletion to cloud
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (url) {
                try {
                    await fetch(url, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                        body: JSON.stringify({ action: "delete_region", sap_region_code: String(regCode).trim() })
                    });
                } catch (e) { console.warn(e); }
            }

            renderAdminKpisAndSummaries();
            renderAdminZoneProgress();
            renderAdminProductionMatrix();
            renderAdminRegionsTable(document.getElementById('admin-region-search')?.value || '');
            showToast(`Cleared data for ${r.region_name}`);
        }

        async function deleteAllCampaignData() {
            const promptVal = prompt("Type 'DELETE ALL' in capital letters to reset all portal & Google Sheet data:");
            if (promptVal !== 'DELETE ALL') {
                alert("Reset cancelled.");
                return;
            }

            store = {};
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));

            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (url) {
                try {
                    await fetch(url, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                        body: JSON.stringify({ action: "reset_all" })
                    });
                } catch (e) { console.warn(e); }
            }

            renderAdminKpisAndSummaries();
            renderAdminZoneProgress();
            renderAdminProductionMatrix();
            renderAdminRegionsTable();
            showToast("⚠️ All data has been completely reset!");
        }

        // -------------------------------------------------------------
        // GOOGLE APPS SCRIPT CLOUD INTEGRATION
        // -------------------------------------------------------------
        function toggleCloudSettings() {
            const drawer = document.getElementById('cloud-settings-drawer');
            drawer.classList.toggle('hidden');
            document.getElementById('custom-cloud-url-input').value = cloudApiUrl || DEFAULT_CLOUD_URL;
        }

        function saveCloudUrlSetting() {
            const url = document.getElementById('custom-cloud-url-input').value.trim();
            cloudApiUrl = url || DEFAULT_CLOUD_URL;
            localStorage.setItem('EXIUM_CLOUD_URL', cloudApiUrl);
            alert("Google Apps Script URL saved successfully!");
        }

        async function testCloudConnection() {
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            showCloudTestResult("Testing connection...", "text-indigo-300 bg-indigo-950/60");
            try {
                const res = await fetch(url + (url.includes('?') ? '&' : '?') + 'action=ping&_t=' + Date.now());
                const data = await res.json();
                if (data && data.status === 'success') {
                    showCloudTestResult("✅ Connected! Google Sheet is active.", "text-emerald-300 bg-emerald-950/60");
                } else {
                    showCloudTestResult("⚠️ Response received but status was not success.", "text-amber-300 bg-amber-950/60");
                }
            } catch (err) {
                showCloudTestResult(`❌ Connection test failed: ${err.message}`, "text-rose-300 bg-rose-950/60");
            }
        }

        function showCloudTestResult(msg, className) {
            const el = document.getElementById('cloud-test-result');
            if (!el) return;
            el.innerHTML = msg;
            el.className = `text-xs font-bold p-2.5 rounded-xl border border-slate-700 ${className}`;
            el.classList.remove('hidden');
        }

        async function syncTerritoryToCloud(terrCode, terrData) {
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (!url || !terrData || !terrCode) return;

            try {
                await fetch(url, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({
                        action: "save_territory",
                        sap_territory_code: String(terrCode).trim(),
                        data: terrData
                    })
                });
            } catch (err) {
                console.warn("[Cloud Sync Error]:", err);
            }
        }

        async function pushAllStoredDataToGoogleSheet() {
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            const allCodes = Object.keys(store);
            if (allCodes.length === 0) {
                alert("There are no stored submissions to push yet.");
                return;
            }

            if (!confirm(`Push all ${allCodes.length} territory submissions to your Google Sheet now?`)) return;

            showToast("⏳ Pushing submissions to Google Sheet...");

            try {
                await fetch(url, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({ action: "save_batch", batch: store })
                });
                showToast(`✅ Successfully pushed ${allCodes.length} territories!`);
                alert(`✅ Successfully pushed ${allCodes.length} territory submissions to your Google Sheet in Google Drive!`);
            } catch (err) {
                alert(`❌ Failed to push: ${err.message}`);
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
                        reject(new Error("Timeout"));
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
                if (res.ok) json = await res.json();
            } catch (fetchErr) {
                console.warn("[Cloud Fetch Failed, Trying JSONP]:", fetchErr);
            }

            // 2. Fallback to JSONP
            if (!json || json.status !== 'success' || !json.store) {
                try {
                    json = await fetchCloudDataJsonp(url);
                } catch (jsonpErr) {
                    console.warn("[JSONP Failed]:", jsonpErr);
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
                renderAdminKpisAndSummaries();
                renderAdminZoneProgress();
                renderAdminProductionMatrix();
                renderAdminRegionsTable(document.getElementById('admin-region-search')?.value || '');
                if (showFeedback) {
                    showToast(`✅ Synced with Google Sheet! (${populatedCount} active entries)`);
                }
                return { success: true, count: populatedCount };
            } else {
                if (showFeedback) showToast("⚠️ Could not pull cloud data.");
                return { success: false, count: 0 };
            }
        }

        // -------------------------------------------------------------
        // EXCEL EXPORTS (LIVE FETCH FROM CLOUD BEFORE DOWNLOAD)
        // -------------------------------------------------------------
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
                    "Status": (d.c1_doc_name && d.c1_doc_rpl && String(d.c1_doc_rpl).length === 6 && d.c1_m1_sweater && d.c1_m1_size && d.c1_m2_sweater && d.c1_m2_size && d.c1_m3_sweater && d.c1_m3_size && (!d.c1_m4_sweater || (d.c1_m4_sweater && d.c1_m4_size))) ? "Complete" : (d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater ? "In Progress" : "Not Started")
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
                    "Doctor 4 Name": '',
                    "Doctor 4 RPL ID": '',
                    "Sweater 4": '',
                    "Size 4": '',
                    "Status": (d.c2_d1_name && d.c2_d1_rpl && String(d.c2_d1_rpl).length === 6 && d.c2_d1_sweater && d.c2_d1_size && d.c2_d2_name && d.c2_d2_rpl && String(d.c2_d2_rpl).length === 6 && d.c2_d2_sweater && d.c2_d2_size && d.c2_d3_name && d.c2_d3_rpl && String(d.c2_d3_rpl).length === 6 && d.c2_d3_sweater && d.c2_d3_size) ? "Complete" : (d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater || d.c2_d2_name || d.c2_d3_name ? "In Progress" : "Not Started")
                });
            });

            const wb = XLSX.utils.book_new();
            const ws1 = XLSX.utils.json_to_sheet(c1Rows);
            const ws2 = XLSX.utils.json_to_sheet(c2Rows);

            XLSX.utils.book_append_sheet(wb, ws1, "Gyne Core Doctor (Family)");
            XLSX.utils.book_append_sheet(wb, ws2, "Core Doctor Maximization");

            XLSX.writeFile(wb, filename);
            showToast("📥 Excel file downloaded with all live data!");
        }
    </script>
</body>
</html>
"""

final_html = html_template.replace("###B64_LOGO###", b64_logo)
final_html = final_html.replace("###B64_01###", b64_01)
final_html = final_html.replace("###B64_02###", b64_02)
final_html = final_html.replace("###B64_03###", b64_03)
final_html = final_html.replace("###B64_04###", b64_04)
final_html = final_html.replace("###B64_05###", b64_05)
final_html = final_html.replace("###ZONE_OPTIONS###", zone_options)
final_html = final_html.replace("###DEFAULT_CLOUD_URL###", DEFAULT_CLOUD_URL)
final_html = final_html.replace("###REGION_MAP###", json.dumps(region_map))
final_html = final_html.replace("###ALL_TERRITORIES###", json.dumps(territories))
final_html = final_html.replace("###ZONES###", json.dumps(zones))

portal_path = os.path.join(base_dir, "Sweater_Campaign_Portal.html")
index_path = os.path.join(base_dir, "index.html")

with open(portal_path, "w", encoding="utf-8") as f:
    f.write(final_html)

with open(index_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print("Successfully regenerated web app cleanly!")
