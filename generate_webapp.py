import json
import os
import io
import base64
import pandas as pd
from PIL import Image

excel_file = r"G:\Exium\2026\4Q'26\Sweater\FF list.xlsx"
df = pd.read_excel(excel_file)
territories = df.to_dict(orient='records')

region_map = {}
for t in territories:
    reg_code = str(t['SAP Region Code'])
    if reg_code not in region_map:
        region_map[reg_code] = {
            'sap_region_code': reg_code,
            'region_name': t['Region'],
            'zone': t['Zone'],
            'regional_head': t['Regional Head'],
            'territories': []
        }
    region_map[reg_code]['territories'].append({
        'sap_territory_code': str(t['SAP Territory Code']),
        'territory_name': t['Territory']
    })

zones = sorted(list(set(df['Zone'].tolist())))
base_dir = os.path.dirname(excel_file)

def get_image_base64(path, max_dim=1000, quality=85):
    if not os.path.exists(path):
        return ""
    im = Image.open(path)
    im.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    if path.lower().endswith('.png'):
        im.save(buf, format='PNG', optimize=True)
        mime = 'image/png'
    else:
        im.save(buf, format='JPEG', quality=quality, optimize=True)
        mime = 'image/jpeg'
    b64_str = base64.b64encode(buf.getvalue()).decode('utf-8')
    return f"data:{mime};base64,{b64_str}"

print("Encoding images...")
b64_logo = get_image_base64(os.path.join(base_dir, 'Exium MUPS Logo.png'), max_dim=600)
b64_01 = get_image_base64(os.path.join(base_dir, 'Image', '01 (Men).jpeg'), max_dim=1000, quality=85)
b64_02 = get_image_base64(os.path.join(base_dir, 'Image', '02 (Men).jpeg'), max_dim=1000, quality=85)
b64_03 = get_image_base64(os.path.join(base_dir, 'Image', '03 (Men).jpeg'), max_dim=1000, quality=85)
b64_04 = get_image_base64(os.path.join(base_dir, 'Image', '04 (Female).jpeg'), max_dim=1000, quality=85)
b64_05 = get_image_base64(os.path.join(base_dir, 'Image', '05 (Female).jpeg'), max_dim=1000, quality=85)

zone_options = '<option value="">-- Choose Your Zone (35 Zones) --</option>\\n'
for z in zones:
    zone_options += f'                        <option value="{z}">{z}</option>\\n'

DEFAULT_CLOUD_URL = "https://script.google.com/macros/s/AKfycbzEnDTtNiXEAyB5qHqrxLj1RbNytgOJAB_lKjw_VVVd1C8CiaeYU6iTROiJabkyX_-b/exec"

html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Exium MUPS - Sweater for Doctors</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800;900&display=swap');
        * { box-sizing: border-box; }
        html { scroll-behavior: smooth; }
        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #f1f5f9;
            color: #1e293b;
            -webkit-tap-highlight-color: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar { width: 5px; height: 5px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f8fafc; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .sweater-card-img {
            image-rendering: -webkit-optimize-contrast;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .sweater-card-img:active { transform: scale(0.96); }
        @media (hover: hover) {
            .sweater-card-img:hover { transform: scale(1.03); }
        }
        .sticky-territory-banner {
            position: -webkit-sticky;
            position: sticky;
            top: 78px;
            z-index: 30;
        }
        @media (min-width: 640px) {
            .sticky-territory-banner { top: 54px; }
        }
    </style>
</head>
<body class="min-h-screen flex flex-col bg-slate-100 text-slate-800 antialiased">

    <!-- Top Navigation Header -->
    <header class="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-2 sm:py-2.5">
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                
                <!-- Line 1: Logo + Title -->
                <div class="flex items-center justify-between sm:justify-start gap-2.5">
                    <div class="flex items-center gap-2 sm:gap-2.5 min-w-0">
                        <img src="###B64_LOGO###" onerror="this.src='Exium MUPS Logo.png'" alt="Exium MUPS" class="h-7 sm:h-8 md:h-9 w-auto object-contain flex-shrink-0">
                        <div class="border-l-2 border-slate-300 pl-2 sm:pl-2.5 flex items-center gap-1.5 sm:gap-2 min-w-0">
                            <h1 class="text-sm sm:text-base md:text-lg font-black text-slate-900 tracking-tight leading-none whitespace-nowrap">Sweater for Doctors</h1>
                            <span class="text-[10px] sm:text-xs bg-orange-500 text-white font-black px-1.5 sm:px-2 py-0.5 rounded-full leading-none shadow-sm flex-shrink-0">4Q'26</span>
                        </div>
                    </div>

                    <!-- Desktop Buttons -->
                    <div class="hidden sm:flex items-center gap-2">
                        <button onclick="openCatalogModal()" class="px-3.5 py-1.5 bg-orange-50 hover:bg-orange-100 text-orange-800 border border-orange-200 rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm active:scale-95 whitespace-nowrap">
                            <i class="fa-solid fa-vest text-orange-600"></i>
                            <span>Catalogue & Sizes</span>
                        </button>
                        <div id="header-admin-btn-container">
                            <button onclick="openAdminModal()" class="px-3.5 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 transition shadow-sm active:scale-95 whitespace-nowrap">
                                <i class="fa-solid fa-shield-halved text-orange-400"></i>
                                <span>Admin</span>
                            </button>
                        </div>
                    </div>
                </div>

                <!-- Line 2 on Mobile -->
                <div class="flex sm:hidden items-center gap-2 pt-1 border-t border-slate-100">
                    <button onclick="openCatalogModal()" class="flex-1 py-1.5 px-3 bg-orange-50 hover:bg-orange-100 text-orange-800 border border-orange-200 rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 shadow-sm active:scale-95">
                        <i class="fa-solid fa-vest text-orange-600"></i>
                        <span>Catalogue & Sizes</span>
                    </button>
                    <div id="mobile-admin-btn-container" class="flex-shrink-0">
                        <button onclick="openAdminModal()" class="py-1.5 px-3 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold flex items-center gap-1 shadow-sm active:scale-95">
                            <i class="fa-solid fa-shield-halved text-orange-400"></i>
                            <span>Admin</span>
                        </button>
                    </div>
                </div>

            </div>
        </div>
    </header>

    <!-- Floating Toast Notification -->
    <div id="toast-notification" class="fixed bottom-5 right-5 z-50 transform translate-y-10 opacity-0 pointer-events-none transition-all duration-300">
        <div class="bg-slate-900 text-white text-xs sm:text-sm font-bold px-4 py-3 rounded-2xl shadow-2xl border border-slate-700 flex items-center gap-2.5">
            <i class="fa-solid fa-circle-check text-emerald-400 text-base"></i>
            <span id="toast-msg">Territory saved successfully!</span>
        </div>
    </div>

    <!-- Main Content Area -->
    <main class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-3 sm:py-6 flex-1 w-full flex flex-col gap-4 sm:gap-6">
        
        <!-- Global Lock Banner -->
        <div id="login-global-locked-alert" class="hidden bg-amber-500 text-slate-950 font-bold px-4 py-2.5 rounded-2xl shadow-sm text-xs sm:text-sm flex items-center gap-2">
            <i class="fa-solid fa-triangle-exclamation text-base"></i>
            <span><strong>Notice:</strong> Submissions are currently locked by Central Admin. Inputs are in Read-Only mode.</span>
        </div>

        <!-- VIEW 1: REGIONAL LOGIN -->
        <section id="selection-view" class="flex-1 flex items-center justify-center py-4 sm:py-8">
            <div class="bg-white border border-slate-200 rounded-3xl shadow-xl p-5 sm:p-8 max-w-lg w-full">
                
                <div class="text-center space-y-1.5 mb-6">
                    <div class="inline-flex p-3 rounded-2xl bg-orange-50 text-orange-600 mb-1 border border-orange-100">
                        <i class="fa-solid fa-user-shield text-2xl"></i>
                    </div>
                    <h2 class="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">Regional Manager Login</h2>
                    <p class="text-xs sm:text-sm text-slate-500">Select your Zone and Region to access your territory portal</p>
                </div>

                <div class="space-y-4">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                            1. Select Zone
                        </label>
                        <select id="select-zone" onchange="onZoneChanged()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition">
###ZONE_OPTIONS###                        </select>
                    </div>

                    <div>
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                            2. Select Region
                        </label>
                        <select id="select-region" onchange="onRegionChanged()" disabled class="w-full bg-slate-100 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition disabled:opacity-50">
                            <option value="">-- Select Zone First --</option>
                        </select>
                    </div>

                    <div id="rh-info-card" class="hidden bg-slate-50 border border-slate-200 rounded-2xl p-3.5 space-y-1">
                        <span class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Regional Head</span>
                        <div id="rh-name-display" class="text-sm font-black text-slate-900 flex items-center gap-2">
                            <i class="fa-solid fa-user-tie text-orange-500"></i>
                            <span>-</span>
                        </div>
                        <div id="rh-territory-count" class="text-xs text-slate-500 font-medium">
                            Total Territories: <strong>0</strong>
                        </div>
                    </div>

                    <div id="password-section" class="hidden space-y-1.5">
                        <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                            3. Enter Password
                        </label>
                        <div class="relative">
                            <input type="password" id="region-password" onkeydown="handlePasswordKey(event)" placeholder="Enter password..." class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-orange-500 focus:border-orange-500 focus:outline-none transition pr-10">
                            <button type="button" onclick="togglePasswordVisibility('region-password', this)" class="absolute right-3 top-2.5 text-slate-400 hover:text-slate-600">
                                <i class="fa-regular fa-eye"></i>
                            </button>
                        </div>
                    </div>

                    <div id="unlock-btn-container" class="hidden">
                        <button onclick="unlockRegion()" class="w-full py-3 bg-orange-500 hover:bg-orange-600 text-white font-black text-sm rounded-xl shadow-lg shadow-orange-500/25 transition flex items-center justify-center gap-2 active:scale-98">
                            <i class="fa-solid fa-right-to-bracket"></i>
                            <span>Login</span>
                        </button>
                    </div>

                </div>

            </div>
        </section>

        <!-- VIEW 2: REGIONAL MANAGER WORKSPACE -->
        <section id="workspace-view" class="hidden flex-col gap-4 sm:gap-6">
            
            <div class="bg-white border border-slate-200 rounded-3xl p-4 sm:p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div class="space-y-1">
                    <div class="flex items-center gap-2">
                        <span id="banner-zone" class="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-orange-100 text-orange-800 border border-orange-200">Zone</span>
                        <span id="banner-region" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200">SAP: 00000</span>
                    </div>
                    <h2 id="banner-rh" class="text-lg sm:text-2xl font-black text-slate-900 tracking-tight">Region Name</h2>
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

            <!-- Layout -->
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-4 sm:gap-6">
                
                <!-- Left Nav -->
                <div class="lg:col-span-3 space-y-3">
                    <div class="block lg:hidden bg-white border border-slate-200 rounded-2xl p-3 shadow-sm">
                        <label class="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">Switch Territory</label>
                        <select id="mobile-territory-select" onchange="selectTerritoryTab(parseInt(this.value))" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none">
                        </select>
                    </div>

                    <div class="hidden lg:flex flex-col bg-white border border-slate-200 rounded-3xl p-4 shadow-sm space-y-3">
                        <div class="flex items-center justify-between border-b border-slate-100 pb-2">
                            <h3 class="text-xs font-black uppercase tracking-wider text-slate-500">Territories</h3>
                            <span id="region-progress-badge" class="text-[10px] font-black px-2 py-0.5 rounded-full bg-slate-100 text-slate-700">0/0 Done</span>
                        </div>
                        <div id="desktop-territory-list" class="space-y-1.5 max-h-[600px] overflow-y-auto custom-scrollbar pr-1">
                        </div>
                    </div>
                </div>

                <!-- Right Form Entries -->
                <div class="lg:col-span-9 space-y-4 sm:space-y-5">

                    <!-- STICKY ACTIVE TERRITORY BAR -->
                    <div id="active-territory-banner-card" class="sticky-territory-banner bg-slate-900/95 backdrop-blur-md border border-slate-800 rounded-2xl p-3 sm:px-5 sm:py-3.5 flex items-center justify-between shadow-xl text-white">
                        <div class="flex items-center gap-2.5 sm:gap-3 min-w-0">
                            <div class="w-8 h-8 sm:w-9 sm:h-9 rounded-xl bg-orange-500 text-slate-950 font-black flex items-center justify-center text-xs sm:text-sm flex-shrink-0 shadow">
                                <i class="fa-solid fa-map-pin"></i>
                            </div>
                            <div class="min-w-0">
                                <div class="flex items-center gap-2">
                                    <span class="text-[9px] sm:text-[10px] font-bold uppercase tracking-wider text-orange-400 block leading-none">Active Territory</span>
                                    <span id="current-territory-status" class="text-[9px] sm:text-[10px] font-bold px-2 py-0.2 rounded-full bg-white/10 text-slate-200 border border-white/20">Not Started</span>
                                </div>
                                <h3 id="current-territory-title" class="text-xs sm:text-base font-black text-white truncate leading-tight mt-0.5">Territory Name</h3>
                                <p id="current-territory-code" class="text-[10px] sm:text-[11px] text-slate-400 font-mono truncate">SAP Code: -</p>
                            </div>
                        </div>
                        <div class="flex items-center gap-2 flex-shrink-0">
                            <button onclick="saveCurrentTerritoryClick()" class="px-3.5 sm:px-4 py-2 bg-emerald-500 hover:bg-emerald-400 text-slate-950 text-xs font-black rounded-xl flex items-center gap-1.5 shadow-lg shadow-emerald-500/20 transition active:scale-95">
                                <i class="fa-solid fa-floppy-disk text-xs"></i>
                                <span>Save</span>
                            </button>
                        </div>
                    </div>

                    <!-- Locked Notice -->
                    <div id="territory-locked-notice" class="hidden bg-rose-50 border border-rose-200 rounded-2xl p-3 sm:p-3.5 text-xs text-rose-800 flex items-center gap-2">
                        <i class="fa-solid fa-lock text-rose-600 text-base flex-shrink-0"></i>
                        <div>
                            <strong>Locked by Admin:</strong> Submissions are locked in view-only mode.
                        </div>
                    </div>

                    <!-- CAMPAIGN 1 -->
                    <div class="bg-white border-2 border-teal-500/60 rounded-3xl shadow-sm overflow-hidden">
                        <div class="bg-gradient-to-r from-teal-700 via-teal-800 to-emerald-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-7 h-7 sm:w-8 sm:h-8 rounded-xl bg-white text-teal-800 flex items-center justify-center font-black text-xs sm:text-sm flex-shrink-0 shadow-sm mt-0.5 sm:mt-0">1</div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Gyne Core Doctor Development (Family Package)</h4>
                                    <p class="text-[10px] sm:text-xs text-teal-100 mt-0.5 leading-tight">1 Doctor per Territory &bull; 4 Sweaters for Family</p>
                                </div>
                            </div>
                            <div class="self-start sm:self-auto pl-9 sm:pl-0">
                                <span class="text-[10px] sm:text-xs font-black bg-teal-950/80 text-teal-200 border border-teal-400/40 px-2.5 py-0.5 rounded-full inline-block whitespace-nowrap shadow-sm">4 Sweaters Total</span>
                            </div>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
                            <div class="bg-teal-50/70 rounded-2xl p-3 sm:p-4 border border-teal-200">
                                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                    <div>
                                        <label class="block text-xs font-bold text-teal-950 mb-1">Doctor Name <span class="text-rose-500">*</span></label>
                                        <input type="text" id="c1_doc_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Enter Gynecologist / Doctor Name..." class="w-full bg-white border border-teal-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 focus:outline-none transition">
                                    </div>
                                    <div>
                                        <div class="flex items-center justify-between mb-1">
                                            <label class="text-xs font-bold text-teal-950">Doctor RPL ID (6 Digits) <span class="text-rose-500">*</span></label>
                                            <span id="c1_doc_rpl_badge" class="text-[10px] font-bold text-slate-400">6 digits</span>
                                        </div>
                                        <input type="text" inputmode="numeric" maxlength="6" id="c1_doc_rpl" oninput="onRplInput(this, 'c1_doc_rpl_badge')" onchange="onRplInput(this, 'c1_doc_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c1_doc_rpl_badge'), 50)" onblur="onRplInput(this, 'c1_doc_rpl_badge')" placeholder="e.g. 104523" class="w-full bg-white border border-teal-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 font-mono font-bold placeholder-slate-400 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 focus:outline-none transition tracking-wider">
                                        <div id="c1_doc_rpl_dup_msg" class="hidden"></div>
                                    </div>
                                </div>
                            </div>

                            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4">
                                <!-- 1 -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">1</span> Sweater 1 (Family Member)</span>
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
                                <!-- 2 -->
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
                                <!-- 3 -->
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
                                <!-- 4 -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">4</span> Sweater 4 (Family Member)</span>
                                        <span id="c1_m4_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
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

                    <!-- CAMPAIGN 2 -->
                    <div class="bg-white border-2 border-purple-500/60 rounded-3xl shadow-sm overflow-hidden">
                        <div class="bg-gradient-to-r from-purple-700 via-purple-800 to-indigo-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-7 h-7 sm:w-8 sm:h-8 rounded-xl bg-white text-purple-800 flex items-center justify-center font-black text-xs sm:text-sm flex-shrink-0 shadow-sm mt-0.5 sm:mt-0">2</div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Core Doctor Maximization</h4>
                                    <p class="text-[10px] sm:text-xs text-purple-100 mt-0.5 leading-tight">4 Doctors per Territory &bull; 1 Sweater Each</p>
                                </div>
                            </div>
                            <div class="self-start sm:self-auto pl-9 sm:pl-0">
                                <span class="text-[10px] sm:text-xs font-black bg-purple-950/80 text-purple-200 border border-purple-400/40 px-2.5 py-0.5 rounded-full inline-block whitespace-nowrap shadow-sm">4 Sweaters Total</span>
                            </div>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
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
                                            <input type="text" id="c2_d1_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Enter Doctor 1 Name..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                        </div>
                                        <div>
                                            <div class="flex items-center justify-between"><label class="text-[10px] font-bold text-purple-950">Doctor 1 RPL ID <span class="text-rose-500">*</span></label><span id="c2_d1_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span></div>
                                            <input type="text" inputmode="numeric" maxlength="6" id="c2_d1_rpl" oninput="onRplInput(this, 'c2_d1_rpl_badge')" onchange="onRplInput(this, 'c2_d1_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d1_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d1_rpl_badge')" placeholder="6-digit RPL ID..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-purple-500 tracking-wider">
                                            <div id="c2_d1_rpl_dup_msg" class="hidden"></div>
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
                                            <input type="text" id="c2_d2_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Enter Doctor 2 Name..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                        </div>
                                        <div>
                                            <div class="flex items-center justify-between"><label class="text-[10px] font-bold text-purple-950">Doctor 2 RPL ID <span class="text-rose-500">*</span></label><span id="c2_d2_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span></div>
                                            <input type="text" inputmode="numeric" maxlength="6" id="c2_d2_rpl" oninput="onRplInput(this, 'c2_d2_rpl_badge')" onchange="onRplInput(this, 'c2_d2_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d2_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d2_rpl_badge')" placeholder="6-digit RPL ID..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-purple-500 tracking-wider">
                                            <div id="c2_d2_rpl_dup_msg" class="hidden"></div>
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
                                <!-- Doc 3 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3.5 space-y-2.5">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">3</span> Doctor 3</span>
                                        <span id="c2_d3_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="space-y-2">
                                        <div>
                                            <label class="text-[10px] font-bold text-purple-950">Doctor 3 Name <span class="text-rose-500">*</span></label>
                                            <input type="text" id="c2_d3_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Enter Doctor 3 Name..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                        </div>
                                        <div>
                                            <div class="flex items-center justify-between"><label class="text-[10px] font-bold text-purple-950">Doctor 3 RPL ID <span class="text-rose-500">*</span></label><span id="c2_d3_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span></div>
                                            <input type="text" inputmode="numeric" maxlength="6" id="c2_d3_rpl" oninput="onRplInput(this, 'c2_d3_rpl_badge')" onchange="onRplInput(this, 'c2_d3_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d3_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d3_rpl_badge')" placeholder="6-digit RPL ID..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-purple-500 tracking-wider">
                                            <div id="c2_d3_rpl_dup_msg" class="hidden"></div>
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
                                <!-- Doc 4 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3.5 space-y-2.5">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">4</span> Doctor 4</span>
                                        <span id="c2_d4_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="space-y-2">
                                        <div>
                                            <label class="text-[10px] font-bold text-purple-950">Doctor 4 Name <span class="text-rose-500">*</span></label>
                                            <input type="text" id="c2_d4_name" oninput="onDataChanged(); validateAllRplFields();" placeholder="Enter Doctor 4 Name..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                        </div>
                                        <div>
                                            <div class="flex items-center justify-between"><label class="text-[10px] font-bold text-purple-950">Doctor 4 RPL ID <span class="text-rose-500">*</span></label><span id="c2_d4_rpl_badge" class="text-[9px] font-bold text-slate-400">6 digits</span></div>
                                            <input type="text" inputmode="numeric" maxlength="6" id="c2_d4_rpl" oninput="onRplInput(this, 'c2_d4_rpl_badge')" onchange="onRplInput(this, 'c2_d4_rpl_badge')" onpaste="setTimeout(() => onRplInput(this, 'c2_d4_rpl_badge'), 50)" onblur="onRplInput(this, 'c2_d4_rpl_badge')" placeholder="6-digit RPL ID..." class="w-full mt-0.5 bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-purple-500 tracking-wider">
                                            <div id="c2_d4_rpl_dup_msg" class="hidden"></div>
                                        </div>
                                    </div>
                                    <div class="flex gap-2.5 sm:gap-3 items-center pt-2 border-t border-purple-200/80">
                                        <div id="c2_d4_img_preview" onclick="zoomSlotImage('c2_d4_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group"><i class="fa-solid fa-shirt text-lg text-slate-300"></i></div>
                                        <div class="flex-1 space-y-1.5 min-w-0">
                                            <div>
                                                <label class="text-[10px] font-bold text-slate-500">Sweater Option</label>
                                                <select id="c2_d4_sweater" onchange="onSweaterSelectChange('c2_d4', this.value)" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1.5 text-xs text-slate-900 font-semibold focus:outline-none focus:border-purple-500">
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
                                                <select id="c2_d4_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500"><option value="">-- Size --</option></select>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Bottom Nav -->
                    <div class="flex items-center justify-between bg-white border border-slate-200 rounded-2xl p-3 sm:p-4 shadow-sm gap-2">
                        <button onclick="saveCurrentTerritoryClick()" class="px-4 sm:px-6 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-sm transition active:scale-95">
                            <i class="fa-solid fa-floppy-disk"></i>
                            <span>Save Territory</span>
                        </button>
                        <div class="flex items-center gap-2">
                            <button onclick="navigateTerritory(-1)" class="px-3 sm:px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold flex items-center gap-1 border border-slate-200 transition active:scale-95">
                                <i class="fa-solid fa-chevron-left"></i> <span class="hidden sm:inline">Previous</span>
                            </button>
                            <button onclick="navigateTerritory(1)" class="px-4 sm:px-5 py-2.5 bg-orange-500 hover:bg-orange-600 text-white font-black rounded-xl text-xs flex items-center gap-1 transition shadow-sm active:scale-95">
                                <span>Next</span> <i class="fa-solid fa-chevron-right"></i>
                            </button>
                        </div>
                    </div>

                </div>

            </div>

        </section>

    </main>

    <!-- LIGHTBOX MODAL -->
    <div id="image-lightbox-modal" class="fixed inset-0 z-[100] bg-slate-950/90 backdrop-blur-md hidden flex items-center justify-center p-3 sm:p-4" onclick="closeImageLightbox()">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[92vh] overflow-y-auto custom-scrollbar relative" onclick="event.stopPropagation()">
            <button onclick="closeImageLightbox()" class="absolute top-3.5 right-3.5 z-20 w-8 h-8 sm:w-9 sm:h-9 rounded-full bg-slate-100 text-slate-700 hover:bg-slate-200 flex items-center justify-center transition border border-slate-200 shadow-sm">
                <i class="fa-solid fa-xmark text-sm sm:text-base"></i>
            </button>
            <div class="p-4 sm:p-6 flex flex-col sm:flex-row gap-4 sm:gap-6 items-center">
                <div class="w-full sm:w-1/2 aspect-[3/4] bg-slate-50 rounded-2xl overflow-hidden border border-slate-200 shadow-sm flex-shrink-0 relative">
                    <img id="lightbox-img" src="" alt="Sweater HD" class="w-full h-full object-cover">
                    <span id="lightbox-code-badge" class="absolute top-3 left-3 bg-slate-900 text-white text-xs font-black px-2.5 py-1 rounded-lg shadow">01</span>
                </div>
                <div class="w-full sm:w-1/2 space-y-3">
                    <div>
                        <span id="lightbox-gender" class="text-[10px] font-bold uppercase tracking-wider text-orange-700 bg-orange-50 px-2 py-0.5 rounded-full border border-orange-200">Men's</span>
                        <h3 id="lightbox-title" class="text-sm sm:text-base font-black text-slate-900 mt-1 leading-snug">Sweater Name</h3>
                        <p id="lightbox-color" class="text-xs text-slate-500 mt-0.5">Color</p>
                    </div>
                    <div class="bg-slate-50 rounded-2xl p-3 border border-slate-200 space-y-2 text-xs">
                        <div class="flex justify-between border-b border-slate-200/60 pb-1.5"><span class="text-slate-500">Supplier:</span><span class="font-bold text-slate-800">Richman / Lubnan</span></div>
                        <div class="flex justify-between border-b border-slate-200/60 pb-1.5"><span class="text-slate-500">Available Sizes:</span><span id="lightbox-sizes" class="font-black text-orange-600">S, M, L, XL, XXL</span></div>
                    </div>
                    <button onclick="closeImageLightbox()" class="w-full py-2.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold rounded-xl transition shadow">Close Preview</button>
                </div>
            </div>
        </div>
    </div>

    <!-- CATALOGUE MODAL -->
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

                <!-- ============================================== -->
                <!-- DETAILED SIZE & MEASUREMENT SPECIFICATIONS    -->
                <!-- ============================================== -->
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

                    <!-- Measurement Guide Note -->
                    <div class="bg-amber-50 border border-amber-200 rounded-2xl p-3.5 flex items-start gap-3 text-xs text-amber-900">
                        <i class="fa-solid fa-circle-info text-amber-600 text-base mt-0.5 flex-shrink-0"></i>
                        <div class="space-y-1">
                            <p class="font-bold">How to Select the Perfect Size for Doctors:</p>
                            <ul class="list-disc pl-4 space-y-0.5 text-[11px] text-amber-800">
                                <li><strong>Chest / Bust:</strong> Measure around the fullest part of the chest. If in-between sizes, choose the larger size for comfortable layering.</li>
                                <li><strong>Body Length:</strong> Measured from the highest point of the shoulder down to the bottom rib hem.</li>
                                <li><strong>Fabric & Quality:</strong> Crafted with high-grade ultra-soft acrylic yarn by <strong>Lubnan Trade Consortium Ltd. (Richman / Lubnan)</strong> for long-lasting winter comfort.</li>
                            </ul>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- ADMIN MODAL -->
    <div id="admin-modal" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-2 sm:p-4 lg:p-6" onclick="closeAdminModal()">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-6xl w-full max-h-[95vh] flex flex-col overflow-hidden" onclick="event.stopPropagation()">
            
            <div class="px-4 sm:px-6 py-3.5 border-b border-slate-200 flex items-center justify-between bg-slate-900 text-white flex-shrink-0">
                <div class="flex items-center gap-3">
                    <div class="w-9 h-9 rounded-2xl bg-orange-500 text-slate-950 flex items-center justify-center font-black text-base shadow">
                        <i class="fa-solid fa-shield-halved"></i>
                    </div>
                    <div>
                        <h3 class="font-black text-sm sm:text-base tracking-tight">Central Admin Control & Live Dashboard</h3>
                        <p class="text-[10px] sm:text-xs text-slate-400">Total 35 Zones &bull; 252 Regions &bull; 1,856 Territories</p>
                    </div>
                </div>
                <button onclick="closeAdminModal()" class="w-8 h-8 rounded-full bg-slate-800 border border-slate-700 text-slate-300 hover:text-white flex items-center justify-center transition">
                    <i class="fa-solid fa-xmark text-sm"></i>
                </button>
            </div>

            <div class="p-4 sm:p-6 overflow-y-auto custom-scrollbar space-y-6 flex-1 bg-slate-50/50">
                <div id="admin-auth-view" class="max-w-sm mx-auto py-10 text-center space-y-4 bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
                    <div class="w-14 h-14 rounded-2xl bg-orange-50 text-orange-600 flex items-center justify-center mx-auto text-2xl border border-orange-200 shadow-inner">
                        <i class="fa-solid fa-key"></i>
                    </div>
                    <div>
                        <h4 class="text-base font-black text-slate-900">Admin Authentication</h4>
                        <p class="text-xs text-slate-500">Enter master password to access central dashboard</p>
                    </div>
                    <div class="space-y-3">
                        <input type="password" id="admin-pass-input" onkeydown="handleAdminPasswordKey(event)" placeholder="Enter Admin Password..." class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs text-slate-900 text-center font-bold focus:ring-2 focus:ring-orange-500 focus:outline-none">
                        <button onclick="verifyAdminPassword()" class="w-full py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-black shadow transition active:scale-98">Unlock Dashboard</button>
                        <div id="admin-auth-err" class="hidden text-xs text-rose-600 font-bold">Incorrect password! Try again.</div>
                    </div>
                </div>

                <div id="admin-dashboard-view" class="hidden space-y-6">
                    <div class="bg-white p-4 rounded-3xl border border-slate-200 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-3">
                        <div class="flex flex-wrap items-center gap-2">
                            <button onclick="pullCloudData(true)" class="px-3.5 py-2 bg-blue-600 hover:bg-blue-500 text-white text-xs font-black rounded-xl flex items-center gap-1.5 shadow-sm transition active:scale-95">
                                <i class="fa-solid fa-rotate"></i>
                                <span>Pull Cloud Data</span>
                            </button>
                            <button onclick="exportMasterExcelFromAdmin()" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black rounded-xl flex items-center gap-1.5 shadow-md shadow-emerald-600/20 transition active:scale-95">
                                <i class="fa-solid fa-file-excel text-sm"></i>
                                <span>Export Live Master Excel (.xlsx)</span>
                            </button>
                            <button onclick="toggleCloudSettings()" class="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 text-xs font-bold rounded-xl flex items-center gap-1.5 transition">
                                <i class="fa-brands fa-google-drive text-amber-600"></i>
                                <span>Google Drive / Cloud API</span>
                            </button>
                        </div>

                        <div class="flex items-center gap-2 self-start md:self-auto">
                            <button onclick="deleteAllCampaignData()" class="px-3.5 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-black rounded-xl flex items-center gap-1.5 transition shadow-sm active:scale-95" title="Reset and Delete All Submissions">
                                <i class="fa-solid fa-triangle-exclamation"></i>
                                <span>Delete All Data</span>
                            </button>
                            <div class="flex items-center gap-1.5 bg-slate-50 border border-slate-200 px-3 py-1.5 rounded-xl text-xs">
                                <span class="text-slate-500 font-medium">Access:</span>
                                <button id="toggle-global-access-btn" onclick="toggleGlobalSubmissionsAccess()" class="font-bold text-emerald-600 hover:underline">Open</button>
                            </div>
                            <button onclick="logoutAdmin()" class="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold rounded-xl transition" title="Logout"><i class="fa-solid fa-lock"></i></button>
                        </div>
                    </div>

                    <!-- Cloud Settings Box -->
                    <div id="admin-cloud-settings-box" class="hidden bg-gradient-to-br from-slate-900 to-slate-800 text-white p-5 rounded-3xl border border-slate-700 shadow-xl space-y-4">
                        <div class="flex items-center justify-between border-b border-slate-700 pb-3">
                            <div class="flex items-center gap-2">
                                <i class="fa-brands fa-google-drive text-orange-400 text-lg"></i>
                                <h4 class="text-sm font-black text-white">Google Drive / Cloud Synchronization Setup</h4>
                            </div>
                            <button onclick="toggleCloudSettings()" class="text-slate-400 hover:text-white text-xs"><i class="fa-solid fa-xmark text-sm"></i></button>
                        </div>
                        <div class="space-y-2">
                            <label class="block text-xs font-bold text-slate-300">Google Apps Script Web App URL (Google Sheets API):</label>
                            <div class="flex flex-col sm:flex-row gap-2">
                                <input type="text" id="custom-cloud-url-input" placeholder="https://script.google.com/macros/s/.../exec" class="flex-1 bg-slate-950 border border-slate-700 rounded-xl px-3 py-2 text-xs text-white font-mono focus:outline-none focus:border-orange-500">
                                <button onclick="saveCloudUrlSetting()" class="px-4 py-2 bg-orange-500 hover:bg-orange-400 text-slate-950 text-xs font-black rounded-xl transition whitespace-nowrap">Save URL</button>
                                <button onclick="testGoogleDriveConnection()" class="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white text-xs font-black rounded-xl transition whitespace-nowrap">Test Connection</button>
                            </div>
                            <div id="cloud-test-result" class="hidden text-xs font-bold p-2 rounded-xl"></div>
                        </div>
                        <div class="pt-3 border-t border-slate-800 flex flex-wrap items-center justify-between gap-3">
                            <div class="flex flex-wrap items-center gap-2">
                                <button onclick="pushAllStoredDataToGoogleSheet()" class="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black shadow transition">
                                    <i class="fa-solid fa-cloud-arrow-up"></i> Push All Stored Data to Google Sheet
                                </button>
                                <button onclick="downloadAllDataJson()" class="px-3 py-2 bg-slate-800 hover:bg-slate-700 border border-slate-600 rounded-xl text-xs font-bold text-slate-200">
                                    <i class="fa-solid fa-download"></i> Backup JSON
                                </button>
                                <label class="px-3 py-2 bg-slate-800 hover:bg-slate-700 border border-slate-600 rounded-xl text-xs font-bold text-slate-200 cursor-pointer">
                                    <i class="fa-solid fa-upload"></i> Restore JSON
                                    <input type="file" id="import-json-file" accept=".json" onchange="importDataJson(event)" class="hidden">
                                </label>
                            </div>
                        </div>
                    </div>

                    <!-- Summary Cards -->
                    <div class="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4">
                        <div class="bg-white border border-slate-200 rounded-3xl p-4 sm:p-5 shadow-sm space-y-2">
                            <div class="flex items-center justify-between text-slate-400"><span class="text-[10px] sm:text-xs font-black uppercase tracking-wider">Total Territories</span><i class="fa-solid fa-map-location-dot text-base text-slate-400"></i></div>
                            <div class="flex items-baseline gap-1.5"><span id="kpi-total-territories" class="text-2xl sm:text-3xl font-black text-slate-900">1,856</span><span class="text-[11px] font-bold text-slate-500">Territories</span></div>
                            <p class="text-[10px] text-slate-400">252 Regions across 35 Zones</p>
                        </div>
                        <div class="bg-white border-2 border-emerald-500/40 rounded-3xl p-4 sm:p-5 shadow-sm space-y-2">
                            <div class="flex items-center justify-between text-emerald-600"><span class="text-[10px] sm:text-xs font-black uppercase tracking-wider">Completed</span><i class="fa-solid fa-circle-check text-base"></i></div>
                            <div class="flex items-baseline gap-1.5"><span id="kpi-completed-count" class="text-2xl sm:text-3xl font-black text-emerald-600">0</span><span id="kpi-completed-pct" class="text-xs font-black text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded-md">0%</span></div>
                            <div class="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden"><div id="kpi-completed-bar" class="bg-emerald-500 h-full rounded-full transition-all duration-500" style="width: 0%"></div></div>
                        </div>
                        <div class="bg-white border border-amber-300 rounded-3xl p-4 sm:p-5 shadow-sm space-y-2">
                            <div class="flex items-center justify-between text-amber-600"><span class="text-[10px] sm:text-xs font-black uppercase tracking-wider">In Progress</span><i class="fa-solid fa-spinner text-base"></i></div>
                            <div class="flex items-baseline gap-1.5"><span id="kpi-inprogress-count" class="text-2xl sm:text-3xl font-black text-amber-600">0</span><span id="kpi-inprogress-pct" class="text-xs font-black text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded-md">0%</span></div>
                            <div class="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden"><div id="kpi-inprogress-bar" class="bg-amber-500 h-full rounded-full transition-all duration-500" style="width: 0%"></div></div>
                        </div>
                        <div class="bg-white border border-slate-200 rounded-3xl p-4 sm:p-5 shadow-sm space-y-2">
                            <div class="flex items-center justify-between text-slate-400"><span class="text-[10px] sm:text-xs font-black uppercase tracking-wider">Remaining</span><i class="fa-solid fa-clock text-base text-slate-400"></i></div>
                            <div class="flex items-baseline gap-1.5"><span id="kpi-pending-count" class="text-2xl sm:text-3xl font-black text-slate-700">1,856</span><span id="kpi-pending-pct" class="text-xs font-black text-slate-600 bg-slate-100 px-1.5 py-0.5 rounded-md">100%</span></div>
                            <div class="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden"><div id="kpi-pending-bar" class="bg-slate-400 h-full rounded-full transition-all duration-500" style="width: 100%"></div></div>
                        </div>
                    </div>

                    <!-- Zone Progress -->
                    <div class="bg-white rounded-3xl border border-slate-200 p-4 sm:p-5 shadow-sm space-y-3">
                        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
                            <div><h4 class="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-900">Zone-Wise Completion Progress (35 Zones)</h4><p class="text-[10px] sm:text-xs text-slate-500">Live submission status across all pharmaceutical sales zones</p></div>
                            <span id="zone-summary-badge" class="text-xs font-black px-2.5 py-1 bg-orange-100 text-orange-800 rounded-xl self-start sm:self-auto">0 / 35 Zones Finished</span>
                        </div>
                        <div class="overflow-x-auto max-h-[300px] overflow-y-auto custom-scrollbar">
                            <table class="w-full text-xs text-left border-collapse">
                                <thead class="bg-slate-100 text-slate-700 uppercase text-[10px] sticky top-0 border-b border-slate-200">
                                    <tr><th class="p-2.5">Zone Name</th><th class="p-2.5 text-center">Regions</th><th class="p-2.5 text-center">Territories</th><th class="p-2.5 text-center">Done</th><th class="p-2.5 text-center">Remaining</th><th class="p-2.5">Progress Bar</th><th class="p-2.5 text-right">Status</th></tr>
                                </thead>
                                <tbody id="admin-zone-table-body" class="divide-y divide-slate-100 font-semibold"></tbody>
                            </table>
                        </div>
                    </div>

                    <!-- Matrix -->
                    <div class="bg-white rounded-3xl border border-slate-200 p-4 sm:p-5 shadow-sm space-y-3">
                        <div class="border-b border-slate-100 pb-3"><h4 class="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-900">Live Production Requirement Breakdown by Size</h4><p class="text-[10px] sm:text-xs text-slate-500">Exact sweater counts for Lubnan Trade Consortium Ltd. order generation</p></div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-xs text-center border-collapse">
                                <thead class="bg-slate-900 text-white text-[11px] font-bold">
                                    <tr><th class="p-2.5 text-left">Item Code & Name</th><th class="p-2.5">XS</th><th class="p-2.5">S</th><th class="p-2.5">M</th><th class="p-2.5">L</th><th class="p-2.5">XL</th><th class="p-2.5">XXL</th><th class="p-2.5 bg-orange-600 text-white">Total</th></tr>
                                </thead>
                                <tbody id="admin-matrix-body" class="divide-y divide-slate-100 font-semibold"></tbody>
                            </table>
                        </div>
                    </div>

                    <!-- 252 Regions Table -->
                    <div class="bg-white rounded-3xl border border-slate-200 p-4 sm:p-5 shadow-sm space-y-3">
                        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
                            <div><h4 class="text-xs sm:text-sm font-black uppercase tracking-wider text-slate-900">All 252 Regions Status & Actions</h4><p class="text-[10px] sm:text-xs text-slate-500">Manage individual regional submissions and lock states</p></div>
                            <div class="flex flex-wrap items-center gap-2">
                                <div class="inline-flex bg-slate-100 p-1 rounded-xl text-xs font-bold">
                                    <button onclick="setAdminRegionFilter('all')" id="rf-tab-all" class="px-2.5 py-1 rounded-lg bg-white text-slate-900 shadow-sm">All (252)</button>
                                    <button onclick="setAdminRegionFilter('complete')" id="rf-tab-complete" class="px-2.5 py-1 rounded-lg text-slate-600 hover:text-slate-900">Complete</button>
                                    <button onclick="setAdminRegionFilter('progress')" id="rf-tab-progress" class="px-2.5 py-1 rounded-lg text-slate-600 hover:text-slate-900">In Progress</button>
                                    <button onclick="setAdminRegionFilter('pending')" id="rf-tab-pending" class="px-2.5 py-1 rounded-lg text-slate-600 hover:text-slate-900">Remaining</button>
                                </div>
                                <input type="text" id="admin-region-search" oninput="filterAdminRegions(this.value)" placeholder="Search Region, Zone, or RH..." class="bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-orange-500 w-48 sm:w-56">
                            </div>
                        </div>
                        <div class="overflow-x-auto max-h-[360px] overflow-y-auto custom-scrollbar">
                            <table class="w-full text-xs text-left border-collapse">
                                <thead class="bg-slate-100 text-slate-600 uppercase text-[10px] sticky top-0 border-b border-slate-200">
                                    <tr><th class="p-2.5">SAP Code</th><th class="p-2.5">Region</th><th class="p-2.5">Zone</th><th class="p-2.5">Regional Head</th><th class="p-2.5">Progress</th><th class="p-2.5">Lock State</th><th class="p-2.5 text-right">Actions</th></tr>
                                </thead>
                                <tbody id="admin-regions-table-body" class="divide-y divide-slate-100 font-medium"></tbody>
                            </table>
                        </div>
                    </div>

                </div>

            </div>
        </div>
    </div>

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

        let store = JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{}');
        let regionLocks = JSON.parse(localStorage.getItem('EXIUM_REGION_LOCKS') || '{}');
        let isGlobalAccessOpen = JSON.parse(localStorage.getItem('EXIUM_GLOBAL_ACCESS') || 'true');
        let cloudApiUrl = localStorage.getItem('EXIUM_CLOUD_URL') || DEFAULT_CLOUD_URL;

        let currentRegionCode = null;
        let activeTerritoryIndex = 0;
        let isAdminLoggedIn = false;
        let currentAdminRegionFilter = 'all';
        let autoSyncTimeout = null;

        window.addEventListener('DOMContentLoaded', async () => {
            populateZoneDropdown();
            checkGlobalLockBanner();
            pullCloudData(false);

            const savedSession = JSON.parse(localStorage.getItem('EXIUM_ACTIVE_SESSION') || 'null');
            if (savedSession && savedSession.region_code && REGION_MAP[savedSession.region_code]) {
                unlockRegion(savedSession.region_code, true);
                if (typeof savedSession.territory_idx === 'number') {
                    selectTerritoryTab(savedSession.territory_idx, false);
                }
            }
        });

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
            const unlockBtn = document.getElementById('unlock-btn-container');

            if (rhCard) rhCard.classList.add('hidden');
            if (passCard) passCard.classList.add('hidden');
            if (unlockBtn) unlockBtn.classList.add('hidden');

            if (!zone) {
                regSel.innerHTML = '<option value="">-- Select Zone First --</option>';
                regSel.disabled = true;
                return;
            }

            const matchingRegions = Object.values(REGION_MAP).filter(r => r.zone === zone);
            matchingRegions.sort((a, b) => a.region_name.localeCompare(b.region_name));

            regSel.innerHTML = '<option value="">-- Choose Region (' + matchingRegions.length + ' Regions) --</option>';
            matchingRegions.forEach(r => {
                const opt = document.createElement('option');
                opt.value = r.sap_region_code;
                opt.textContent = `${r.region_name} (${r.sap_region_code})`;
                regSel.appendChild(opt);
            });
            regSel.disabled = false;
        }

        function onRegionChanged() {
            const regCode = document.getElementById('select-region').value;
            const rhCard = document.getElementById('rh-info-card');
            const passCard = document.getElementById('password-section');
            const unlockBtn = document.getElementById('unlock-btn-container');
            const passInput = document.getElementById('region-password');

            if (passInput) passInput.value = '';

            if (!regCode || !REGION_MAP[regCode]) {
                if (rhCard) rhCard.classList.add('hidden');
                if (passCard) passCard.classList.add('hidden');
                if (unlockBtn) unlockBtn.classList.add('hidden');
                return;
            }

            const r = REGION_MAP[regCode];
            document.getElementById('rh-name-display').querySelector('span').textContent = r.regional_head;
            document.getElementById('rh-territory-count').innerHTML = `Total Territories: <strong>${r.territories.length}</strong>`;

            if (rhCard) rhCard.classList.remove('hidden');
            if (passCard) passCard.classList.remove('hidden');
            if (unlockBtn) unlockBtn.classList.remove('hidden');
            if (passInput) passInput.focus();
        }

        function handlePasswordKey(e) {
            if (e.key === 'Enter') unlockRegion();
        }

        function unlockRegion(bypassCode = null, isRestoringSession = false) {
            const code = bypassCode || document.getElementById('select-region').value;
            const passInput = document.getElementById('region-password');
            const pass = passInput ? passInput.value.trim() : '';

            if (!bypassCode && pass !== code && pass !== 'Exium MUPS') {
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

            document.getElementById('banner-zone').textContent = r.zone;
            document.getElementById('banner-region').textContent = `SAP: ${r.sap_region_code}`;
            document.getElementById('banner-rh').textContent = `Region: ${r.region_name} (${r.regional_head})`;

            renderTerritoryTabs();
            if (!isRestoringSession) {
                selectTerritoryTab(0, true);
            }
        }

        function exitRegionWorkspace() {
            onDataChanged();
            localStorage.removeItem('EXIUM_ACTIVE_SESSION');
            currentRegionCode = null;
            document.getElementById('workspace-view').classList.add('hidden');
            document.getElementById('selection-view').classList.remove('hidden');
            const passInput = document.getElementById('region-password');
            if (passInput) passInput.value = '';
        }

        function renderTerritoryTabs() {
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
                mobSelect.appendChild(mobOpt);

                const btn = document.createElement('button');
                btn.type = 'button';
                btn.onclick = () => selectTerritoryTab(idx);
                btn.id = `terr-tab-btn-${idx}`;
                btn.className = `w-full text-left p-3 rounded-2xl text-xs font-bold transition flex items-center justify-between border ${
                    idx === activeTerritoryIndex 
                    ? 'bg-orange-500 text-white border-orange-500 shadow-md' 
                    : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'
                }`;

                btn.innerHTML = `
                    <div class="truncate pr-1 min-w-0">
                        <div class="truncate font-black text-xs">${t.territory_name}</div>
                        <div class="text-[10px] ${idx === activeTerritoryIndex ? 'text-orange-100' : 'text-slate-400'} font-mono">SAP: ${t.sap_territory_code}</div>
                    </div>
                    <span class="text-[9px] font-black px-1.5 py-0.5 rounded-full flex-shrink-0 ${
                        status === 'Complete' ? (idx === activeTerritoryIndex ? 'bg-white text-slate-950' : 'bg-emerald-100 text-emerald-800 border border-emerald-300') :
                        status === 'In Progress' ? (idx === activeTerritoryIndex ? 'bg-white text-slate-950' : 'bg-amber-100 text-amber-800 border border-amber-300') :
                        (idx === activeTerritoryIndex ? 'bg-orange-600 text-white' : 'bg-slate-200 text-slate-600')
                    }">${status === 'Complete' ? '✓ Complete' : status}</span>
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
            statusBadge.className = `text-[9px] sm:text-[10px] font-bold px-2 py-0.2 rounded-full ${
                status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-bold' :
                'bg-white/10 text-slate-200 border border-white/20'
            }`;

            const isLocked = isRegionLocked();
            const lockedNotice = document.getElementById('territory-locked-notice');
            if (isLocked) {
                lockedNotice.classList.remove('hidden');
            } else {
                lockedNotice.classList.add('hidden');
            }

            const c1DocInput = document.getElementById('c1_doc_name');
            c1DocInput.value = d.c1_doc_name || '';
            c1DocInput.disabled = isLocked;

            const c1DocRpl = document.getElementById('c1_doc_rpl');
            c1DocRpl.value = d.c1_doc_rpl || '';
            c1DocRpl.disabled = isLocked;
            updateRplBadgeState(c1DocRpl, 'c1_doc_rpl_badge');

            ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                const sw = d[`c1_${m}_sweater`] || '';
                const sz = d[`c1_${m}_size`] || '';
                const swSel = document.getElementById(`c1_${m}_sweater`);
                const szSel = document.getElementById(`c1_${m}_size`);
                
                swSel.value = sw;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c1_${m}_sweater`, `c1_${m}_size`, sz);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c1_${m}_img_preview`, sw);
                updateSweaterSlotIndicator(`c1_${m}`);
            });

            ['d1', 'd2', 'd3', 'd4'].forEach(d_item => {
                const dNameInput = document.getElementById(`c2_${d_item}_name`);
                dNameInput.value = d[`c2_${d_item}_name`] || '';
                dNameInput.disabled = isLocked;

                const dRplInput = document.getElementById(`c2_${d_item}_rpl`);
                dRplInput.value = d[`c2_${d_item}_rpl`] || '';
                dRplInput.disabled = isLocked;
                updateRplBadgeState(dRplInput, `c2_${d_item}_rpl_badge`);

                const sw = d[`c2_${d_item}_sweater`] || '';
                const sz = d[`c2_${d_item}_size`] || '';
                const swSel = document.getElementById(`c2_${d_item}_sweater`);
                const szSel = document.getElementById(`c2_${d_item}_size`);

                swSel.value = sw;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c2_${d_item}_sweater`, `c2_${d_item}_size`, sz);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c2_${d_item}_img_preview`, sw);
                updateSweaterSlotIndicator(`c2_${d_item}`);
            });

            r.territories.forEach((_, tabIdx) => {
                const btn = document.getElementById(`terr-tab-btn-${tabIdx}`);
                if (btn) {
                    if (tabIdx === idx) {
                        btn.className = 'w-full text-left p-3 rounded-2xl text-xs font-bold transition flex items-center justify-between border bg-orange-500 text-white border-orange-500 shadow-md';
                    } else {
                        btn.className = 'w-full text-left p-3 rounded-2xl text-xs font-bold transition flex items-center justify-between border bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200';
                    }
                }
            });

            if (shouldScroll) {
                const bannerEl = document.getElementById('active-territory-banner-card');
                if (bannerEl) {
                    bannerEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            }
        }

        // MAP FOR FAST TERRITORY LOOKUP
        const ALL_TERRITORIES_MAP = {};
        ALL_TERRITORIES.forEach(t => {
            ALL_TERRITORIES_MAP[String(t['SAP Territory Code']).trim()] = t;
        });

        // DOCTOR RPL DUPLICATE DETECTOR
        function findDoctorRplDuplicate(rplVal, currentFieldId) {
            if (!rplVal) return null;
            rplVal = String(rplVal).replace(/[^0-9]/g, '').trim();
            if (rplVal.length !== 6) return null;

            if (!currentRegionCode) return null;
            const r = REGION_MAP[currentRegionCode];
            const currentTerritory = r.territories[activeTerritoryIndex];
            const currentTerrCode = String(currentTerritory.sap_territory_code).trim();

            // 1. Check other input fields on the active territory form in DOM
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
                            zone: currentTerritory.zone,
                            campaignLabel: f.label,
                            doctorName: docName || '(Doctor name not entered yet)'
                        };
                    }
                }
            }

            // 2. Check store across all territories in the entire campaign
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
                    // Skip current field of current territory
                    if (isCurrentTerr && slot.fieldId === currentFieldId) continue;
                    // Current territory other slots were already checked in active DOM above
                    if (isCurrentTerr) continue;

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
            statusBadge.className = `text-[9px] sm:text-[10px] font-bold px-2 py-0.2 rounded-full ${
                status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-bold' :
                'bg-white/10 text-slate-200 border border-white/20'
            }`;

            const isLocked = isRegionLocked();
            const lockedNotice = document.getElementById('territory-locked-notice');
            if (isLocked) {
                lockedNotice.classList.remove('hidden');
            } else {
                lockedNotice.classList.add('hidden');
            }

            const c1DocInput = document.getElementById('c1_doc_name');
            c1DocInput.value = d.c1_doc_name || '';
            c1DocInput.disabled = isLocked;

            const c1DocRpl = document.getElementById('c1_doc_rpl');
            c1DocRpl.value = d.c1_doc_rpl || '';
            c1DocRpl.disabled = isLocked;
            updateRplBadgeState(c1DocRpl, 'c1_doc_rpl_badge');

            ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                const sw = d[`c1_${m}_sweater`] || '';
                const sz = d[`c1_${m}_size`] || '';
                const swSel = document.getElementById(`c1_${m}_sweater`);
                const szSel = document.getElementById(`c1_${m}_size`);
                
                swSel.value = sw;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c1_${m}_sweater`, `c1_${m}_size`, sz);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c1_${m}_img_preview`, sw);
                updateSweaterSlotIndicator(`c1_${m}`);
            });

            ['d1', 'd2', 'd3', 'd4'].forEach(d_item => {
                const dNameInput = document.getElementById(`c2_${d_item}_name`);
                dNameInput.value = d[`c2_${d_item}_name`] || '';
                dNameInput.disabled = isLocked;

                const dRplInput = document.getElementById(`c2_${d_item}_rpl`);
                dRplInput.value = d[`c2_${d_item}_rpl`] || '';
                dRplInput.disabled = isLocked;
                updateRplBadgeState(dRplInput, `c2_${d_item}_rpl_badge`);

                const sw = d[`c2_${d_item}_sweater`] || '';
                const sz = d[`c2_${d_item}_size`] || '';
                const swSel = document.getElementById(`c2_${d_item}_sweater`);
                const szSel = document.getElementById(`c2_${d_item}_size`);

                swSel.value = sw;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c2_${d_item}_sweater`, `c2_${d_item}_size`, sz);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c2_${d_item}_img_preview`, sw);
                updateSweaterSlotIndicator(`c2_${d_item}`);
            });

            r.territories.forEach((_, tabIdx) => {
                const btn = document.getElementById(`terr-tab-btn-${tabIdx}`);
                if (btn) {
                    if (tabIdx === idx) {
                        btn.className = 'w-full text-left p-3 rounded-2xl text-xs font-bold transition flex items-center justify-between border bg-orange-500 text-white border-orange-500 shadow-md';
                    } else {
                        btn.className = 'w-full text-left p-3 rounded-2xl text-xs font-bold transition flex items-center justify-between border bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200';
                    }
                }
            });

            if (shouldScroll) {
                const bannerEl = document.getElementById('active-territory-banner-card');
                if (bannerEl) {
                    bannerEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
                }
            }
        }

        function onRplInput(inputEl, badgeId) {
            inputEl.value = inputEl.value.replace(/[^0-9]/g, '').slice(0, 6);
            updateRplBadgeState(inputEl, badgeId);
            onDataChanged();
        }

        function updateRplBadgeState(inputEl, badgeId) {
            const val = inputEl.value || '';
            const badge = document.getElementById(badgeId);
            if (!badge) return;

            if (val.length === 0) {
                badge.textContent = "6 digits";
                badge.className = "text-[9px] sm:text-[10px] font-bold text-slate-400";
            } else if (val.length < 6) {
                badge.textContent = `${val.length}/6 digits`;
                badge.className = "text-[9px] sm:text-[10px] font-black text-amber-600";
            } else if (val.length === 6) {
                badge.innerHTML = '<i class="fa-solid fa-check text-emerald-600"></i> Valid 6-Digit';
                badge.className = "text-[9px] sm:text-[10px] font-black text-emerald-600";
            }
        }

        function updateSweaterSlotIndicator(slotPrefix) {
            const sw = document.getElementById(`${slotPrefix}_sweater`)?.value || '';
            const sz = document.getElementById(`${slotPrefix}_size`)?.value || '';
            const badge = document.getElementById(`${slotPrefix}_check_badge`);
            if (!badge) return;

            if (sw && sz) {
                badge.innerHTML = `<span class="bg-emerald-100 text-emerald-800 border border-emerald-300 text-[10px] font-black px-2 py-0.5 rounded-full flex items-center gap-1 shadow-sm"><i class="fa-solid fa-circle-check text-emerald-600"></i> Complete</span>`;
            } else if (sw || sz) {
                badge.innerHTML = `<span class="bg-amber-100 text-amber-800 border border-amber-300 text-[10px] font-bold px-1.5 py-0.5 rounded-full flex items-center gap-1"><i class="fa-solid fa-clock text-amber-600"></i> Incomplete</span>`;
            } else {
                badge.innerHTML = `<span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span>`;
            }
        }

        function triggerAutoSync() {
            if (autoSyncTimeout) clearTimeout(autoSyncTimeout);
            autoSyncTimeout = setTimeout(() => {
                if (currentRegionCode && REGION_MAP[currentRegionCode]) {
                    const r = REGION_MAP[currentRegionCode];
                    const t = r.territories[activeTerritoryIndex];
                    if (t) {
                        const terrCode = String(t.sap_territory_code);
                        syncTerritoryToCloud(terrCode, store[terrCode]);
                    }
                }
            }, 1200);
        }

        function onDataChanged() {
            if (isRegionLocked() || !currentRegionCode) return;

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
                c2_d4_size: document.getElementById('c2_d4_size').value,
            };

            store[terrCode] = terrData;
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));
                validateAllRplFields();

            ['c1_m1', 'c1_m2', 'c1_m3', 'c1_m4', 'c2_d1', 'c2_d2', 'c2_d3', 'c2_d4'].forEach(p => updateSweaterSlotIndicator(p));

            const status = getTerritoryStatus(terrData);
            const statusBadge = document.getElementById('current-territory-status');
            statusBadge.textContent = status;
            statusBadge.className = `text-[9px] sm:text-[10px] font-bold px-2 py-0.2 rounded-full ${
                status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-bold' :
                'bg-white/10 text-slate-200 border border-white/20'
            }`;

            let completedCount = 0;
            r.territories.forEach(ter => {
                if (getTerritoryStatus(store[String(ter.sap_territory_code)]) === 'Complete') completedCount++;
            });
            document.getElementById('region-progress-badge').textContent = `${completedCount}/${r.territories.length} Done`;
            
            triggerAutoSync();
        }

        
        // VALIDATE TERRITORY COMPLETENESS BEFORE SAVING
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
            showToast(`📥 Excel file downloaded with all live data!`);
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

print("Successfully regenerated web app with async cloud delete and live Excel export!")
