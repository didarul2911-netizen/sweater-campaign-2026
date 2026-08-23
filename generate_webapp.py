import json
import os
import io
import base64
import pandas as pd
from PIL import Image

# Load FF list
df = pd.read_excel('FF list.xlsx')
territories = df.to_dict(orient='records')

# Build Region Master Map
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

# Helper to encode images as Base64 Data URIs
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

print("Encoding images to Base64...")
b64_logo = get_image_base64('Exium MUPS Logo.png', max_dim=600)
b64_01 = get_image_base64('Image/01 (Men).jpeg', max_dim=1000, quality=85)
b64_02 = get_image_base64('Image/02 (Men).jpeg', max_dim=1000, quality=85)
b64_03 = get_image_base64('Image/03 (Men).jpeg', max_dim=1000, quality=85)
b64_04 = get_image_base64('Image/04 (Female).jpeg', max_dim=1000, quality=85)
b64_05 = get_image_base64('Image/05 (Female).jpeg', max_dim=1000, quality=85)

html_code = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Exium MUPS - Sweater for Doctors</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome 6 -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">
    <!-- SheetJS (xlsx) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800;900&display=swap');
        * {{
            box-sizing: border-box;
        }}
        html {{
            scroll-behavior: smooth;
        }}
        body {{
            font-family: 'Plus Jakarta Sans', sans-serif;
            background-color: #f1f5f9;
            color: #1e293b;
            -webkit-tap-highlight-color: transparent;
        }}
        .custom-scrollbar::-webkit-scrollbar {{
            width: 5px;
            height: 5px;
        }}
        .custom-scrollbar::-webkit-scrollbar-track {{
            background: #f8fafc;
        }}
        .custom-scrollbar::-webkit-scrollbar-thumb {{
            background: #cbd5e1;
            border-radius: 4px;
        }}
        .sweater-card-img {{
            image-rendering: -webkit-optimize-contrast;
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }}
        .sweater-card-img:active {{
            transform: scale(0.96);
        }}
        @media (hover: hover) {{
            .sweater-card-img:hover {{
                transform: scale(1.03);
            }}
        }}
        .sticky-territory-banner {{
            position: -webkit-sticky;
            position: sticky;
            top: 75px;
            z-index: 30;
        }}
        @media (min-width: 640px) {{
            .sticky-territory-banner {{
                top: 54px;
            }}
        }}
    </style>
</head>
<body class="min-h-screen flex flex-col bg-slate-100 text-slate-800 antialiased">

    <!-- Top Navigation Header -->
    <header class="bg-white border-b border-slate-200 shadow-sm sticky top-0 z-40">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-2 sm:py-2.5">
            
            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                
                <!-- Line 1 on Mobile: Logo + Sweater for Doctors in 1 clean line -->
                <div class="flex items-center justify-between sm:justify-start gap-2.5">
                    <div class="flex items-center gap-2 sm:gap-2.5 min-w-0">
                        <img src="{b64_logo}" onerror="this.src='Exium MUPS Logo.png'" alt="Exium MUPS" class="h-7 sm:h-8 md:h-9 w-auto object-contain flex-shrink-0">
                        <div class="border-l-2 border-slate-300 pl-2 sm:pl-2.5 flex items-center gap-1.5 sm:gap-2 min-w-0">
                            <h1 class="text-sm sm:text-base md:text-lg font-black text-slate-900 tracking-tight leading-none whitespace-nowrap">Sweater for Doctors</h1>
                            <span class="text-[10px] sm:text-xs bg-orange-500 text-white font-black px-1.5 sm:px-2 py-0.5 rounded-full leading-none shadow-sm flex-shrink-0">4Q'26</span>
                        </div>
                    </div>

                    <!-- Desktop Action Buttons (Right side) -->
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

                <!-- Line 2 on Mobile: Catalogue & Sizes and Admin Buttons -->
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

    <!-- Main Content Area -->
    <main class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 py-3 sm:py-6 flex-1 w-full flex flex-col gap-4 sm:gap-6">

        <!-- =================================================================== -->
        <!-- VIEW 1: REGION SELECTION & LOGIN SCREEN -->
        <!-- =================================================================== -->
        <section id="selection-view" class="max-w-md mx-auto w-full space-y-4 my-auto py-2 sm:py-6">
            
            <div class="text-center space-y-1 mb-2">
                <div class="w-12 h-12 rounded-2xl bg-gradient-to-tr from-orange-500 to-amber-500 flex items-center justify-center mx-auto shadow-md shadow-orange-500/20 text-white text-xl mb-2 sm:mb-3">
                    <i class="fa-solid fa-user-lock"></i>
                </div>
                <h2 class="text-xl sm:text-2xl font-black text-slate-900 tracking-tight">Regional Manager Login</h2>
                <p class="text-xs text-slate-500">Select Zone & Region, then enter Password.</p>
            </div>

            <!-- Global Locked Alert -->
            <div id="login-global-locked-alert" class="hidden bg-rose-50 border border-rose-200 rounded-2xl p-3 sm:p-4 text-center text-xs text-rose-800 flex items-center justify-center gap-2">
                <i class="fa-solid fa-lock text-rose-600 text-sm"></i>
                <span><strong>Access Closed:</strong> Submissions are currently locked by Central Admin.</span>
            </div>

            <div class="bg-white border border-slate-200 rounded-3xl p-5 sm:p-7 shadow-lg space-y-4">
                
                <!-- 1. Select Zone -->
                <div>
                    <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                        <span class="w-5 h-5 inline-flex items-center justify-center bg-slate-900 text-white font-bold rounded-full text-[10px] mr-1">1</span>
                        Zone
                    </label>
                    <select id="select-zone" onchange="onZoneChanged()" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm text-slate-900 font-semibold focus:ring-2 focus:ring-orange-500 focus:bg-white focus:outline-none transition">
                        <option value="">-- Select Zone --</option>
                    </select>
                </div>

                <!-- 2. Select Region -->
                <div>
                    <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5">
                        <span class="w-5 h-5 inline-flex items-center justify-center bg-slate-900 text-white font-bold rounded-full text-[10px] mr-1">2</span>
                        Region
                    </label>
                    <select id="select-region" onchange="onRegionChanged()" disabled class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm text-slate-900 font-semibold focus:ring-2 focus:ring-orange-500 focus:bg-white focus:outline-none disabled:opacity-50 disabled:cursor-not-allowed transition">
                        <option value="">-- Select Zone First --</option>
                    </select>
                </div>

                <!-- Regional Head Display Card -->
                <div id="rh-info-card" class="hidden bg-orange-50/80 border border-orange-200 rounded-2xl p-3 sm:p-3.5 flex items-center justify-between">
                    <div>
                        <span class="text-[10px] font-bold text-orange-800 uppercase tracking-wider">Regional Head</span>
                        <div id="rh-name-display" class="text-sm font-black text-slate-900 mt-0.5">-</div>
                        <div id="rh-territory-count" class="text-[11px] text-slate-600 mt-0.5">-</div>
                    </div>
                    <div class="w-9 h-9 rounded-xl bg-orange-500 text-white flex items-center justify-center text-sm shadow-sm flex-shrink-0">
                        <i class="fa-solid fa-user-tie"></i>
                    </div>
                </div>

                <!-- 3. Password Input (Supports Enter Key) -->
                <div id="password-section" class="hidden space-y-1.5">
                    <label class="block text-xs font-bold text-slate-700 uppercase tracking-wider">
                        <span class="w-5 h-5 inline-flex items-center justify-center bg-slate-900 text-white font-bold rounded-full text-[10px] mr-1">3</span>
                        Password
                    </label>
                    <div class="relative">
                        <i class="fa-solid fa-lock absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                        <input type="password" id="region-password" onkeydown="handlePasswordKey(event)" placeholder="Enter Password..." class="w-full bg-slate-50 border border-slate-300 rounded-xl pl-9 pr-3.5 py-2.5 text-xs sm:text-sm text-slate-900 font-semibold focus:ring-2 focus:ring-orange-500 focus:bg-white focus:outline-none transition">
                    </div>
                </div>

                <!-- Login Button -->
                <div id="unlock-btn-container" class="hidden pt-1 sm:pt-2">
                    <button onclick="unlockRegion()" class="w-full bg-gradient-to-r from-orange-500 to-amber-500 hover:from-orange-600 hover:to-amber-600 text-white font-black py-3 px-6 rounded-xl shadow-md shadow-orange-500/30 transition transform active:scale-[0.98] flex items-center justify-center gap-2 text-xs sm:text-sm uppercase tracking-wider">
                        <i class="fa-solid fa-arrow-right-to-bracket text-sm"></i>
                        <span>Login</span>
                    </button>
                </div>

            </div>

        </section>


        <!-- =================================================================== -->
        <!-- VIEW 2: REGION WORKSPACE (TERRITORY DATA ENTRY) -->
        <!-- =================================================================== -->
        <section id="workspace-view" class="hidden space-y-3 sm:space-y-5">
            
            <!-- Region Header Banner (Clean - Just Region info & Exit) -->
            <div class="bg-white border border-slate-200 rounded-2xl sm:rounded-3xl p-3 sm:p-4 shadow-sm flex items-center justify-between gap-3">
                <div class="space-y-0.5 min-w-0">
                    <div class="flex items-center gap-1.5 sm:gap-2 flex-wrap">
                        <span id="banner-zone" class="px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-700 text-[11px] sm:text-xs font-bold border border-slate-200">Zone</span>
                        <span id="banner-region" class="px-2.5 py-0.5 rounded-full bg-orange-50 text-orange-700 text-[11px] sm:text-xs font-bold border border-orange-200">Region</span>
                        <span id="banner-locked-status" class="hidden px-2.5 py-0.5 rounded-full bg-rose-100 text-rose-700 text-[11px] sm:text-xs font-bold border border-rose-200 flex items-center gap-1">
                            <i class="fa-solid fa-lock text-[10px]"></i> Locked by Admin
                        </span>
                    </div>
                    <h2 id="banner-rh" class="text-xs sm:text-base font-extrabold text-slate-900 flex items-center gap-1.5 sm:gap-2 truncate">
                        <i class="fa-solid fa-user-tie text-orange-500"></i>
                        <span class="truncate">Regional Head: -</span>
                    </h2>
                </div>

                <div class="flex items-center gap-2 flex-shrink-0">
                    <button onclick="exitRegionWorkspace()" class="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold flex items-center gap-1.5 border border-slate-300 transition active:scale-95 shadow-sm">
                        <i class="fa-solid fa-arrow-right-from-bracket text-slate-500"></i>
                        <span>Exit</span>
                    </button>
                </div>
            </div>

            <!-- Mobile Territory Selector Dropdown -->
            <div class="block lg:hidden bg-white border border-slate-200 rounded-2xl p-3 shadow-sm">
                <div class="flex items-center justify-between mb-1.5">
                    <label class="text-[11px] font-bold text-slate-600 uppercase flex items-center gap-1">
                        <i class="fa-solid fa-location-dot text-orange-500"></i> Select Territory
                    </label>
                    <span id="mobile-progress-badge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-slate-100 text-slate-700">0/0 Done</span>
                </div>
                <select id="mobile-territory-select" onchange="selectTerritoryTab(parseInt(this.value), true)" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-900 focus:ring-2 focus:ring-orange-500 focus:outline-none">
                    <!-- Populated dynamically -->
                </select>
            </div>

            <!-- Workspace Layout Grid -->
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-4 sm:gap-6">

                <!-- Left Column: Desktop Territory Tabs (Hidden on mobile) -->
                <div class="hidden lg:block lg:col-span-3 space-y-3">
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm">
                        <div class="flex items-center justify-between mb-3">
                            <h3 class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                                <i class="fa-solid fa-location-dot text-orange-500"></i> Territories
                            </h3>
                            <span id="region-progress-badge" class="text-[11px] font-bold px-2 py-0.5 rounded-full bg-slate-100 text-slate-700">0/0 Done</span>
                        </div>
                        <div id="territory-tabs-container" class="space-y-1.5 max-h-[580px] overflow-y-auto custom-scrollbar pr-1">
                            <!-- Populated dynamically -->
                        </div>
                    </div>
                </div>

                <!-- Right Column: Campaign Entry Forms -->
                <div class="lg:col-span-9 space-y-4 sm:space-y-5">

                    <!-- FREEZE / STICKY ACTIVE TERRITORY BAR (Always visible during scrolling) -->
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
                            <button onclick="saveCurrentTerritoryClick()" class="px-3 sm:px-4 py-2 bg-emerald-500 hover:bg-emerald-400 text-slate-950 text-xs font-black rounded-xl flex items-center gap-1 shadow-lg shadow-emerald-500/20 transition active:scale-95">
                                <i class="fa-solid fa-floppy-disk text-xs"></i>
                                <span>Save</span>
                            </button>
                        </div>
                    </div>

                    <!-- Locked Notice (if admin locked) -->
                    <div id="territory-locked-notice" class="hidden bg-rose-50 border border-rose-200 rounded-2xl p-3 sm:p-3.5 text-xs text-rose-800 flex items-center gap-2">
                        <i class="fa-solid fa-lock text-rose-600 text-base flex-shrink-0"></i>
                        <div>
                            <strong>Locked by Admin:</strong> Submissions for this region are locked. Inputs are in view-only mode.
                        </div>
                    </div>

                    <!-- ========================================== -->
                    <!-- CAMPAIGN 1: Gyne Core Doctor (Family Pack) -->
                    <!-- Distinct Emerald / Teal Palette with Header Band -->
                    <!-- ========================================== -->
                    <div class="bg-white border-2 border-teal-500/60 rounded-3xl shadow-sm overflow-hidden">
                        
                        <!-- Rich Teal Header Band -->
                        <div class="bg-gradient-to-r from-teal-700 via-teal-800 to-emerald-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-7 h-7 sm:w-8 sm:h-8 rounded-xl bg-white text-teal-800 flex items-center justify-center font-black text-xs sm:text-sm flex-shrink-0 shadow-sm mt-0.5 sm:mt-0">
                                    1
                                </div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Gyne Core Doctor Development (Family Package)</h4>
                                    <p class="text-[10px] sm:text-xs text-teal-100 mt-0.5 leading-tight">1 Doctor per Territory &bull; 4 Sweaters for Family</p>
                                </div>
                            </div>
                            <div class="self-start sm:self-auto pl-9 sm:pl-0">
                                <span class="text-[10px] sm:text-xs font-black bg-teal-950/80 text-teal-200 border border-teal-400/40 px-2.5 py-0.5 rounded-full inline-block whitespace-nowrap shadow-sm">
                                    4 Sweaters Total
                                </span>
                            </div>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
                            <!-- Doctor Name Input (Soft Mint Background) -->
                            <div class="bg-teal-50/70 rounded-2xl p-3 sm:p-3.5 border border-teal-200">
                                <label class="block text-xs font-bold text-teal-950 mb-1">
                                    Doctor Name <span class="text-rose-500">*</span>
                                </label>
                                <input type="text" id="c1_doc_name" oninput="onDataChanged()" placeholder="Enter Gynecologist / Doctor Name..." class="w-full bg-white border border-teal-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:ring-2 focus:ring-teal-500 focus:border-teal-500 focus:outline-none transition">
                            </div>

                            <!-- 4 Sweaters Grid -->
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4">
                                
                                <!-- Sweater 1 -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">1</span>
                                        Sweater 1 (Family Member)
                                    </span>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m1_img_preview" onclick="zoomSlotImage('c1_m1_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c1_m1_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Sweater 2 -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">2</span>
                                        Sweater 2 (Family Member)
                                    </span>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m2_img_preview" onclick="zoomSlotImage('c1_m2_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c1_m2_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Sweater 3 -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">3</span>
                                        Sweater 3 (Family Member)
                                    </span>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m3_img_preview" onclick="zoomSlotImage('c1_m3_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c1_m3_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Sweater 4 -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">4</span>
                                        Sweater 4 (Family Member)
                                    </span>
                                    <div class="flex gap-2.5 sm:gap-3 items-center">
                                        <div id="c1_m4_img_preview" onclick="zoomSlotImage('c1_m4_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c1_m4_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-teal-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                            </div>
                        </div>
                    </div>


                    <!-- ========================================== -->
                    <!-- CAMPAIGN 2: Core Doctor Maximization       -->
                    <!-- Distinct Royal Purple / Indigo Palette with Header Band -->
                    <!-- ========================================== -->
                    <div class="bg-white border-2 border-purple-500/60 rounded-3xl shadow-sm overflow-hidden">
                        
                        <!-- Rich Purple Header Band -->
                        <div class="bg-gradient-to-r from-purple-700 via-purple-800 to-indigo-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-7 h-7 sm:w-8 sm:h-8 rounded-xl bg-white text-purple-800 flex items-center justify-center font-black text-xs sm:text-sm flex-shrink-0 shadow-sm mt-0.5 sm:mt-0">
                                    2
                                </div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Core Doctor Maximization</h4>
                                    <p class="text-[10px] sm:text-xs text-purple-100 mt-0.5 leading-tight">4 Doctors per Territory &bull; 1 Sweater Each</p>
                                </div>
                            </div>
                            <div class="self-start sm:self-auto pl-9 sm:pl-0">
                                <span class="text-[10px] sm:text-xs font-black bg-purple-950/80 text-purple-200 border border-purple-400/40 px-2.5 py-0.5 rounded-full inline-block whitespace-nowrap shadow-sm">
                                    4 Sweaters Total
                                </span>
                            </div>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
                            <!-- 4 Doctors Grid -->
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4">
                                
                                <!-- Doctor 1 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">1</span>
                                        Doctor 1 Name <span class="text-rose-500">*</span>
                                    </span>
                                    <input type="text" id="c2_d1_name" oninput="onDataChanged()" placeholder="Enter Doctor 1 Name..." class="w-full bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                    
                                    <div class="flex gap-2.5 sm:gap-3 items-center pt-1 border-t border-purple-200/80">
                                        <div id="c2_d1_img_preview" onclick="zoomSlotImage('c2_d1_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c2_d1_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Doctor 2 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">2</span>
                                        Doctor 2 Name <span class="text-rose-500">*</span>
                                    </span>
                                    <input type="text" id="c2_d2_name" oninput="onDataChanged()" placeholder="Enter Doctor 2 Name..." class="w-full bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                    
                                    <div class="flex gap-2.5 sm:gap-3 items-center pt-1 border-t border-purple-200/80">
                                        <div id="c2_d2_img_preview" onclick="zoomSlotImage('c2_d2_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c2_d2_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Doctor 3 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">3</span>
                                        Doctor 3 Name <span class="text-rose-500">*</span>
                                    </span>
                                    <input type="text" id="c2_d3_name" oninput="onDataChanged()" placeholder="Enter Doctor 3 Name..." class="w-full bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                    
                                    <div class="flex gap-2.5 sm:gap-3 items-center pt-1 border-t border-purple-200/80">
                                        <div id="c2_d3_img_preview" onclick="zoomSlotImage('c2_d3_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c2_d3_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <!-- Doctor 4 -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3 space-y-2">
                                    <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5">
                                        <span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">4</span>
                                        Doctor 4 Name <span class="text-rose-500">*</span>
                                    </span>
                                    <input type="text" id="c2_d4_name" oninput="onDataChanged()" placeholder="Enter Doctor 4 Name..." class="w-full bg-white border border-purple-300 rounded-xl px-3 py-1.5 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:border-purple-500">
                                    
                                    <div class="flex gap-2.5 sm:gap-3 items-center pt-1 border-t border-purple-200/80">
                                        <div id="c2_d4_img_preview" onclick="zoomSlotImage('c2_d4_sweater')" class="sweater-card-img w-16 h-20 sm:w-20 sm:h-24 rounded-xl bg-white border border-slate-300 overflow-hidden flex-shrink-0 flex items-center justify-center text-slate-400 cursor-pointer shadow-sm relative group">
                                            <i class="fa-solid fa-shirt text-lg text-slate-300"></i>
                                        </div>
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
                                                <select id="c2_d4_size" onchange="onDataChanged()" class="w-full mt-0.5 bg-white border border-slate-300 rounded-lg px-2 py-1 text-xs text-slate-900 font-black focus:outline-none focus:border-purple-500">
                                                    <option value="">-- Size --</option>
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                </div>

                            </div>
                        </div>
                    </div>

                    <!-- Bottom Nav Actions -->
                    <div class="flex items-center justify-between bg-white border border-slate-200 rounded-2xl p-3 sm:p-4 shadow-sm gap-2">
                        <button onclick="saveCurrentTerritoryClick()" class="px-3.5 sm:px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black flex items-center gap-1.5 shadow-sm transition active:scale-95">
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


    <!-- =================================================================== -->
    <!-- MODAL: HD SWEATER ZOOM LIGHTBOX (Z-INDEX 100) -->
    <!-- =================================================================== -->
    <div id="image-lightbox-modal" class="fixed inset-0 z-[100] bg-slate-950/90 backdrop-blur-md hidden flex items-center justify-center p-3 sm:p-4" onclick="closeImageLightbox()">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-2xl w-full max-h-[92vh] overflow-y-auto custom-scrollbar relative" onclick="event.stopPropagation()">
            
            <button onclick="closeImageLightbox()" class="absolute top-3.5 right-3.5 z-20 w-8 h-8 sm:w-9 sm:h-9 rounded-full bg-slate-100 text-slate-700 hover:bg-slate-200 flex items-center justify-center transition border border-slate-200 shadow-sm">
                <i class="fa-solid fa-xmark text-sm sm:text-base"></i>
            </button>

            <div class="p-4 sm:p-6 flex flex-col sm:flex-row gap-4 sm:gap-6 items-center">
                <!-- Large HD Image -->
                <div class="w-full sm:w-1/2 aspect-[3/4] bg-slate-50 rounded-2xl overflow-hidden border border-slate-200 shadow-sm flex-shrink-0 relative">
                    <img id="lightbox-img" src="" alt="Sweater HD" class="w-full h-full object-cover" style="image-rendering: -webkit-optimize-contrast;">
                    <span id="lightbox-code-badge" class="absolute top-3 left-3 bg-slate-900 text-white text-xs font-black px-2.5 py-1 rounded-lg shadow">
                        01
                    </span>
                </div>

                <!-- Product Specifications -->
                <div class="w-full sm:w-1/2 space-y-3">
                    <div>
                        <span id="lightbox-gender" class="text-[10px] font-bold uppercase tracking-wider text-orange-700 bg-orange-50 px-2 py-0.5 rounded-full border border-orange-200">Men's</span>
                        <h3 id="lightbox-title" class="text-sm sm:text-base font-black text-slate-900 mt-1 leading-snug">Men's Sleeveless V-Neck Sweater</h3>
                        <p id="lightbox-color" class="text-xs text-slate-500 mt-0.5">Solid Ash / Grey Textured</p>
                    </div>

                    <div class="bg-slate-50 rounded-xl p-2.5 sm:p-3 border border-slate-200 space-y-0.5 text-xs">
                        <div class="text-[10px] font-bold text-slate-500 uppercase">Available Sizes</div>
                        <div id="lightbox-sizes" class="font-black text-teal-700 text-sm">S, M, L, XL, XXL</div>
                    </div>

                    <div class="bg-slate-50 rounded-xl p-2.5 sm:p-3 border border-slate-200 space-y-1 text-xs">
                        <div class="text-[10px] font-bold text-slate-500 uppercase">Size Measurements (CM)</div>
                        <p id="lightbox-measurements" class="text-[11px] text-slate-700 leading-relaxed font-mono whitespace-pre-line">-</p>
                    </div>

                    <button onclick="closeImageLightbox()" class="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold py-2.5 rounded-xl text-xs transition shadow-sm active:scale-95">
                        Close Preview
                    </button>
                </div>
            </div>

        </div>
    </div>


    <!-- =================================================================== -->
    <!-- MODAL: CATALOGUE & SIZE SPECIFICATIONS (Z-INDEX 50) -->
    <!-- =================================================================== -->
    <div id="catalog-modal" class="fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-md hidden flex items-center justify-center p-3 sm:p-4">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-5xl w-full max-h-[92vh] flex flex-col overflow-hidden text-slate-800">
            <div class="p-3.5 sm:p-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
                <div class="flex items-center gap-2.5 sm:gap-3">
                    <div class="w-8 h-8 sm:w-9 sm:h-9 rounded-xl bg-orange-100 text-orange-600 flex items-center justify-center font-bold text-sm">
                        <i class="fa-solid fa-vest"></i>
                    </div>
                    <div>
                        <h3 class="font-bold text-xs sm:text-sm text-slate-900">Sweater Designs Catalogue & Size Specification</h3>
                        <p class="text-[10px] sm:text-[11px] text-slate-500">Tap on any sweater to enlarge photo & view measurements</p>
                    </div>
                </div>
                <button onclick="closeCatalogModal()" class="text-slate-400 hover:text-slate-700 p-1.5 rounded-lg">
                    <i class="fa-solid fa-xmark text-lg sm:text-xl"></i>
                </button>
            </div>
            
            <div class="p-4 sm:p-6 overflow-y-auto space-y-5 sm:space-y-6 custom-scrollbar">
                <!-- 5 Products in Uniform Grid (Clean numeric triggers: '01', '02', etc.) -->
                <div>
                    <h4 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3 flex items-center gap-1.5">
                        <i class="fa-solid fa-hand-pointer text-orange-500"></i>
                        <span>5 Premium Sweater Choices (Tap to Enlarge)</span>
                    </h4>
                    <div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-3">
                        
                        <!-- 01 -->
                        <div onclick="openImageLightbox('01')" role="button" tabindex="0" class="bg-slate-50 border border-slate-200 hover:border-orange-500 rounded-2xl p-2 sm:p-2.5 text-center flex flex-col cursor-pointer transition transform active:scale-95 shadow-sm">
                            <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden mb-1.5 sm:mb-2 relative border border-slate-200">
                                <img src="{b64_01}" onerror="this.src='Image/01 (Men).jpeg'" alt="Men Grey" class="w-full h-full object-cover" style="image-rendering: -webkit-optimize-contrast;">
                                <span class="absolute top-1.5 left-1.5 bg-slate-900 text-white text-[9px] font-black px-1.5 py-0.5 rounded">01</span>
                                <span class="absolute bottom-1 right-1 bg-white/95 text-slate-900 text-[9px] font-black px-2 py-0.5 rounded-md shadow border border-slate-200 flex items-center gap-1"><i class="fa-solid fa-magnifying-glass-plus text-orange-600"></i> Zoom</span>
                            </div>
                            <span class="text-[11px] sm:text-xs font-bold text-slate-900 truncate">Men's V-Neck (Grey)</span>
                            <span class="text-[9px] sm:text-[10px] text-slate-500 truncate">Solid Ash Textured</span>
                            <span class="text-[9px] sm:text-[10px] text-orange-600 font-bold mt-0.5">Sizes: S - XXL</span>
                        </div>

                        <!-- 02 -->
                        <div onclick="openImageLightbox('02')" role="button" tabindex="0" class="bg-slate-50 border border-slate-200 hover:border-orange-500 rounded-2xl p-2 sm:p-2.5 text-center flex flex-col cursor-pointer transition transform active:scale-95 shadow-sm">
                            <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden mb-1.5 sm:mb-2 relative border border-slate-200">
                                <img src="{b64_02}" onerror="this.src='Image/02 (Men).jpeg'" alt="Men Navy" class="w-full h-full object-cover" style="image-rendering: -webkit-optimize-contrast;">
                                <span class="absolute top-1.5 left-1.5 bg-slate-900 text-white text-[9px] font-black px-1.5 py-0.5 rounded">02</span>
                                <span class="absolute bottom-1 right-1 bg-white/95 text-slate-900 text-[9px] font-black px-2 py-0.5 rounded-md shadow border border-slate-200 flex items-center gap-1"><i class="fa-solid fa-magnifying-glass-plus text-orange-600"></i> Zoom</span>
                            </div>
                            <span class="text-[11px] sm:text-xs font-bold text-slate-900 truncate">Men's V-Neck (Navy)</span>
                            <span class="text-[9px] sm:text-[10px] text-slate-500 truncate">Solid Navy Blue</span>
                            <span class="text-[9px] sm:text-[10px] text-orange-600 font-bold mt-0.5">Sizes: S - XXL</span>
                        </div>

                        <!-- 03 -->
                        <div onclick="openImageLightbox('03')" role="button" tabindex="0" class="bg-slate-50 border border-slate-200 hover:border-orange-500 rounded-2xl p-2 sm:p-2.5 text-center flex flex-col cursor-pointer transition transform active:scale-95 shadow-sm">
                            <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden mb-1.5 sm:mb-2 relative border border-slate-200">
                                <img src="{b64_03}" onerror="this.src='Image/03 (Men).jpeg'" alt="Men Cream" class="w-full h-full object-cover" style="image-rendering: -webkit-optimize-contrast;">
                                <span class="absolute top-1.5 left-1.5 bg-slate-900 text-white text-[9px] font-black px-1.5 py-0.5 rounded">03</span>
                                <span class="absolute bottom-1 right-1 bg-white/95 text-slate-900 text-[9px] font-black px-2 py-0.5 rounded-md shadow border border-slate-200 flex items-center gap-1"><i class="fa-solid fa-magnifying-glass-plus text-orange-600"></i> Zoom</span>
                            </div>
                            <span class="text-[11px] sm:text-xs font-bold text-slate-900 truncate">Men's V-Neck (Cream)</span>
                            <span class="text-[9px] sm:text-[10px] text-slate-500 truncate">Off-White Grid Check</span>
                            <span class="text-[9px] sm:text-[10px] text-orange-600 font-bold mt-0.5">Sizes: S - XXL</span>
                        </div>

                        <!-- 04 -->
                        <div onclick="openImageLightbox('04')" role="button" tabindex="0" class="bg-slate-50 border border-slate-200 hover:border-purple-500 rounded-2xl p-2 sm:p-2.5 text-center flex flex-col cursor-pointer transition transform active:scale-95 shadow-sm">
                            <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden mb-1.5 sm:mb-2 relative border border-slate-200">
                                <img src="{b64_04}" onerror="this.src='Image/04 (Female).jpeg'" alt="Women Cardigan" class="w-full h-full object-cover" style="image-rendering: -webkit-optimize-contrast;">
                                <span class="absolute top-1.5 left-1.5 bg-purple-700 text-white text-[9px] font-black px-1.5 py-0.5 rounded">04</span>
                                <span class="absolute bottom-1 right-1 bg-white/95 text-slate-900 text-[9px] font-black px-2 py-0.5 rounded-md shadow border border-slate-200 flex items-center gap-1"><i class="fa-solid fa-magnifying-glass-plus text-purple-600"></i> Zoom</span>
                            </div>
                            <span class="text-[11px] sm:text-xs font-bold text-slate-900 truncate">Women Cardigan</span>
                            <span class="text-[9px] sm:text-[10px] text-slate-500 truncate">White Check Button</span>
                            <span class="text-[9px] sm:text-[10px] text-purple-700 font-bold mt-0.5">Sizes: XS - XL</span>
                        </div>

                        <!-- 05 -->
                        <div onclick="openImageLightbox('05')" role="button" tabindex="0" class="bg-slate-50 border border-slate-200 hover:border-purple-500 rounded-2xl p-2 sm:p-2.5 text-center flex flex-col cursor-pointer transition transform active:scale-95 shadow-sm">
                            <div class="aspect-[3/4] bg-white rounded-xl overflow-hidden mb-1.5 sm:mb-2 relative border border-slate-200">
                                <img src="{b64_05}" onerror="this.src='Image/05 (Female).jpeg'" alt="Women Semi Long" class="w-full h-full object-cover" style="image-rendering: -webkit-optimize-contrast;">
                                <span class="absolute top-1.5 left-1.5 bg-purple-700 text-white text-[9px] font-black px-1.5 py-0.5 rounded">05</span>
                                <span class="absolute bottom-1 right-1 bg-white/95 text-slate-900 text-[9px] font-black px-2 py-0.5 rounded-md shadow border border-slate-200 flex items-center gap-1"><i class="fa-solid fa-magnifying-glass-plus text-purple-600"></i> Zoom</span>
                            </div>
                            <span class="text-[11px] sm:text-xs font-bold text-slate-900 truncate">Women Semi Long</span>
                            <span class="text-[9px] sm:text-[10px] text-slate-500 truncate">Black Ethnic Border</span>
                            <span class="text-[9px] sm:text-[10px] text-purple-700 font-bold mt-0.5">Sizes: S - XXL</span>
                        </div>

                    </div>
                </div>

                <!-- Size Tables (CM) -->
                <div>
                    <h4 class="text-xs font-bold uppercase tracking-wider text-slate-500 mb-3">Measurement Size Chart (in CM)</h4>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        
                        <!-- Men's Table -->
                        <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 sm:p-4">
                            <div class="text-xs font-bold text-slate-900 mb-2 flex items-center gap-1.5">
                                <i class="fa-solid fa-ruler text-orange-500"></i> Men's Sleeveless V-Neck (Styles 01, 02, 03)
                            </div>
                            <div class="overflow-x-auto">
                                <table class="w-full text-xs text-center border-collapse">
                                    <thead>
                                        <tr class="bg-slate-900 text-white">
                                            <th class="py-1 px-2 text-left">Measurement</th>
                                            <th>S</th><th>M</th><th>L</th><th>XL</th><th>XXL</th>
                                        </tr>
                                    </thead>
                                    <tbody class="divide-y divide-slate-200 bg-white text-slate-700">
                                        <tr><td class="py-1 px-2 text-left font-semibold">Body Length (cm)</td><td>65</td><td>67</td><td>69</td><td>71</td><td>73</td></tr>
                                        <tr><td class="py-1 px-2 text-left font-semibold">1/2 Chest (cm)</td><td>48</td><td>50</td><td>52</td><td>54</td><td>56</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <!-- Women's Table -->
                        <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 sm:p-4 space-y-3">
                            <div class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                                <i class="fa-solid fa-ruler text-purple-600"></i> Women's Cardigans (Styles 04 & 05)
                            </div>
                            <div class="overflow-x-auto space-y-2">
                                <div>
                                    <span class="text-[10px] font-bold text-purple-900">Style 04 (Short Cardigan):</span>
                                    <table class="w-full text-[11px] text-center border-collapse mt-0.5">
                                        <thead>
                                            <tr class="bg-purple-900 text-white">
                                                <th class="py-0.5 px-2 text-left">Point</th><th>XS</th><th>S</th><th>M</th><th>L</th><th>XL</th>
                                            </tr>
                                        </thead>
                                        <tbody class="divide-y divide-slate-200 bg-white text-slate-700">
                                            <tr><td class="py-0.5 px-2 text-left">Length</td><td>60</td><td>62</td><td>64</td><td>66</td><td>68</td></tr>
                                            <tr><td class="py-0.5 px-2 text-left">1/2 Chest</td><td>44</td><td>46</td><td>48</td><td>50</td><td>52</td></tr>
                                            <tr><td class="py-0.5 px-2 text-left">Sleeve</td><td>57</td><td>58</td><td>59</td><td>60</td><td>61</td></tr>
                                        </tbody>
                                    </table>
                                </div>
                                <div>
                                    <span class="text-[10px] font-bold text-purple-900">Style 05 (Semi Long Cardigan):</span>
                                    <table class="w-full text-[11px] text-center border-collapse mt-0.5">
                                        <thead>
                                            <tr class="bg-purple-950 text-white">
                                                <th class="py-0.5 px-2 text-left">Point</th><th>S</th><th>M</th><th>L</th><th>XL</th><th>XXL</th>
                                            </tr>
                                        </thead>
                                        <tbody class="divide-y divide-slate-200 bg-white text-slate-700">
                                            <tr><td class="py-0.5 px-2 text-left">Length</td><td>64</td><td>66</td><td>68</td><td>70</td><td>72</td></tr>
                                            <tr><td class="py-0.5 px-2 text-left">1/2 Chest</td><td>51</td><td>53</td><td>55</td><td>57</td><td>59</td></tr>
                                            <tr><td class="py-0.5 px-2 text-left">Sleeve</td><td>52</td><td>53</td><td>54</td><td>55</td><td>56</td></tr>
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                    </div>
                </div>

            </div>
        </div>
    </div>


    <!-- =================================================================== -->
    <!-- MODAL: ADMIN CONTROL PANEL, LIVE EXPORT & DELETE (Z-INDEX 50) -->
    <!-- =================================================================== -->
    <div id="admin-modal" class="fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-md hidden flex items-center justify-center p-3 sm:p-4">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-5xl w-full max-h-[92vh] flex flex-col overflow-hidden text-slate-800">
            <div class="p-3.5 sm:p-4 border-b border-slate-200 flex items-center justify-between bg-slate-900 text-white">
                <div class="flex items-center gap-2.5 sm:gap-3">
                    <div class="w-8 h-8 sm:w-9 sm:h-9 rounded-xl bg-orange-500 text-slate-950 flex items-center justify-center font-bold text-sm">
                        <i class="fa-solid fa-crown"></i>
                    </div>
                    <div>
                        <h3 class="font-bold text-xs sm:text-sm">Central Admin Control & Master Export</h3>
                        <p class="text-[10px] sm:text-[11px] text-slate-400">Total 252 Regions &bull; 1,856 Territories</p>
                    </div>
                </div>
                <button onclick="closeAdminModal()" class="text-slate-400 hover:text-white p-1 rounded-lg">
                    <i class="fa-solid fa-xmark text-lg sm:text-xl"></i>
                </button>
            </div>
            
            <div class="p-4 sm:p-6 overflow-y-auto space-y-5 sm:space-y-6 custom-scrollbar bg-slate-50/50">
                
                <!-- Admin Auth Box -->
                <div id="admin-login-box" class="max-w-md mx-auto p-5 sm:p-6 bg-white border border-slate-200 rounded-2xl text-center space-y-4 shadow-sm">
                    <div class="w-12 h-12 rounded-2xl bg-orange-50 text-orange-600 flex items-center justify-center mx-auto text-xl border border-orange-200">
                        <i class="fa-solid fa-shield-halved"></i>
                    </div>
                    <div>
                        <h4 class="text-base font-bold text-slate-900">Admin Authentication</h4>
                        <p class="text-xs text-slate-500">Enter Admin Password to continue.</p>
                    </div>
                    <input type="password" id="admin-pass-input" onkeydown="handleAdminPasswordKey(event)" placeholder="Password..." class="w-full bg-slate-50 border border-slate-300 rounded-xl px-4 py-2.5 text-xs text-slate-900 text-center focus:ring-2 focus:ring-orange-500 focus:outline-none">
                    <button onclick="loginAdmin()" class="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold py-2.5 rounded-xl text-xs transition shadow-sm active:scale-95">
                        Authenticate Admin
                    </button>
                </div>

                <!-- Admin Dashboard (When Authenticated) -->
                <div id="admin-dashboard-content" class="hidden space-y-5 sm:space-y-6">
                    
                    <!-- Access Control Bar -->
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-sm">
                        <div>
                            <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Master Controls</span>
                            <div class="text-xs sm:text-sm font-bold text-slate-900 mt-0.5 flex items-center gap-2">
                                <span>Global Submission:</span>
                                <span id="admin-global-access-badge" class="px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">OPEN</span>
                            </div>
                        </div>
                        <div class="flex items-center gap-2 flex-wrap">
                            <button onclick="toggleGlobalAccess()" id="admin-global-toggle-btn" class="px-3 sm:px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white font-bold rounded-xl text-xs transition flex items-center gap-1.5 shadow-sm active:scale-95">
                                <i class="fa-solid fa-lock"></i>
                                <span>Lock Submissions</span>
                            </button>
                            <button onclick="promptDeleteAllData()" class="px-3 sm:px-4 py-2 bg-rose-800 hover:bg-rose-700 text-white font-bold rounded-xl text-xs transition flex items-center gap-1.5 shadow-sm active:scale-95">
                                <i class="fa-solid fa-trash-can"></i>
                                <span>Delete All Submissions</span>
                            </button>
                            <button onclick="exportMasterExcelFromAdmin()" class="px-3 sm:px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-sm transition active:scale-95">
                                <i class="fa-solid fa-file-excel"></i>
                                <span>Download Excel (.xlsx)</span>
                            </button>
                        </div>
                    </div>

                    <!-- KPI Cards -->
                    <div class="grid grid-cols-2 sm:grid-cols-4 gap-2.5 sm:gap-3">
                        <div class="bg-white border border-slate-200 rounded-2xl p-3 sm:p-4 shadow-sm">
                            <span class="text-[10px] font-bold text-slate-500 uppercase">Territories</span>
                            <div class="text-lg sm:text-xl font-black text-slate-900 mt-0.5">1,856</div>
                            <span class="text-[10px] text-slate-500">252 Regions</span>
                        </div>
                        <div class="bg-white border border-slate-200 rounded-2xl p-3 sm:p-4 shadow-sm">
                            <span class="text-[10px] font-bold text-emerald-600 uppercase">Completed</span>
                            <div id="admin-kpi-completed" class="text-lg sm:text-xl font-black text-emerald-600 mt-0.5">0</div>
                            <span id="admin-kpi-completed-pct" class="text-[10px] text-slate-500">0%</span>
                        </div>
                        <div class="bg-white border border-slate-200 rounded-2xl p-3 sm:p-4 shadow-sm">
                            <span class="text-[10px] font-bold text-amber-600 uppercase">In Progress</span>
                            <div id="admin-kpi-inprogress" class="text-lg sm:text-xl font-black text-amber-600 mt-0.5">0</div>
                            <span class="text-[10px] text-slate-500">Partially entered</span>
                        </div>
                        <div class="bg-white border border-slate-200 rounded-2xl p-3 sm:p-4 shadow-sm">
                            <span class="text-[10px] font-bold text-indigo-600 uppercase">Sweaters Count</span>
                            <div id="admin-kpi-sweaters" class="text-lg sm:text-xl font-black text-indigo-600 mt-0.5">0 / 14,848</div>
                            <span class="text-[10px] text-slate-500">Target: 14,848 pcs</span>
                        </div>
                    </div>

                    <!-- Live Procurement Matrix -->
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 space-y-3 shadow-sm">
                        <h4 class="text-xs font-bold uppercase tracking-wider text-slate-700">Live Production Requirement Breakdown</h4>
                        <div id="admin-procurement-matrix-container" class="overflow-x-auto"></div>
                    </div>

                    <!-- Region Progress & Lock/Delete Table -->
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 space-y-3 shadow-sm">
                        <div class="flex items-center justify-between gap-2 flex-wrap">
                            <h4 class="text-xs font-bold uppercase tracking-wider text-slate-700">All 252 Regions Access & Status</h4>
                            <input type="text" id="admin-region-search" oninput="renderAdminRegionTable()" placeholder="Search Region or RH..." class="bg-slate-50 border border-slate-300 rounded-lg px-3 py-1 text-xs text-slate-900">
                        </div>
                        <div class="max-h-72 overflow-y-auto custom-scrollbar border border-slate-200 rounded-xl overflow-x-auto">
                            <table class="w-full text-xs text-left min-w-[500px]">
                                <thead class="bg-slate-100 text-slate-600 uppercase text-[10px] sticky top-0 border-b border-slate-200">
                                    <tr>
                                        <th class="py-2 px-3">SAP Code</th>
                                        <th class="py-2 px-3">Region</th>
                                        <th class="py-2 px-3">Regional Head</th>
                                        <th class="py-2 px-2 text-center">Done</th>
                                        <th class="py-2 px-3 text-right">Actions</th>
                                    </tr>
                                </thead>
                                <tbody id="admin-region-tbody" class="divide-y divide-slate-200 bg-white text-slate-700">
                                    <!-- Populated dynamically -->
                                </tbody>
                            </table>
                        </div>
                    </div>

                </div>

            </div>
        </div>
    </div>


    <!-- =================================================================== -->
    <!-- CONFIRMATION POPUP MODAL FOR DELETE / RESET (Z-INDEX 80) -->
    <!-- =================================================================== -->
    <div id="delete-confirm-modal" class="fixed inset-0 z-[80] bg-slate-900/85 backdrop-blur-sm hidden flex items-center justify-center p-4">
        <div class="bg-white border border-slate-200 rounded-3xl p-6 shadow-2xl max-w-md w-full text-center space-y-4">
            <div class="w-12 h-12 rounded-full bg-rose-100 text-rose-600 flex items-center justify-center mx-auto text-xl">
                <i class="fa-solid fa-triangle-exclamation"></i>
            </div>
            <div>
                <h3 class="text-base font-black text-slate-900">Confirm Delete</h3>
                <p id="delete-confirm-message" class="text-xs text-slate-600 mt-1 leading-relaxed">
                    Are you sure you want to delete this data?
                </p>
            </div>
            <div class="flex items-center justify-center gap-3 pt-2">
                <button onclick="closeDeleteConfirmModal()" class="flex-1 px-4 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition">
                    Cancel
                </button>
                <button id="delete-confirm-btn" class="flex-1 px-4 py-2.5 bg-rose-600 hover:bg-rose-500 text-white rounded-xl text-xs font-bold shadow-md shadow-rose-600/30 transition">
                    Yes, Delete Data
                </button>
            </div>
        </div>
    </div>


    <!-- =================================================================== -->
    <!-- JAVASCRIPT LOGIC & STATE STORE -->
    <!-- =================================================================== -->
    <script>
        const REGION_MAP = {json.dumps(region_map)};
        const ALL_TERRITORIES = {json.dumps(territories)};
        const ZONES = {json.dumps(zones)};

        // Clean Numerical Keys ('01', '02', '03', '04', '05') for 100% Reliable Execution
        const SWEATER_DETAILS = {{
            "01": {{
                code: "01",
                name: "Men's Sleeveless V-Neck Sweater",
                color: "Solid Ash / Grey (Textured Knitted)",
                gender: "Men's",
                sizes: "S, M, L, XL, XXL",
                img: "{b64_01}",
                fallback_img: "Image/01 (Men).jpeg",
                measurements: "Length: S(65cm), M(67cm), L(69cm), XL(71cm), XXL(73cm)\\n1/2 Chest: S(48cm), M(50cm), L(52cm), XL(54cm), XXL(56cm)"
            }},
            "02": {{
                code: "02",
                name: "Men's Sleeveless V-Neck Sweater",
                color: "Solid Navy Blue (Textured Knitted)",
                gender: "Men's",
                sizes: "S, M, L, XL, XXL",
                img: "{b64_02}",
                fallback_img: "Image/02 (Men).jpeg",
                measurements: "Length: S(65cm), M(67cm), L(69cm), XL(71cm), XXL(73cm)\\n1/2 Chest: S(48cm), M(50cm), L(52cm), XL(54cm), XXL(56cm)"
            }},
            "03": {{
                code: "03",
                name: "Men's Sleeveless V-Neck Sweater",
                color: "Off-White / Cream Grid Check",
                gender: "Men's",
                sizes: "S, M, L, XL, XXL",
                img: "{b64_03}",
                fallback_img: "Image/03 (Men).jpeg",
                measurements: "Length: S(65cm), M(67cm), L(69cm), XL(71cm), XXL(73cm)\\n1/2 Chest: S(48cm), M(50cm), L(52cm), XL(54cm), XXL(56cm)"
            }},
            "04": {{
                code: "04",
                name: "Women's Short Cardigan (Button-Up)",
                color: "White & Navy Grid Check",
                gender: "Women's",
                sizes: "XS, S, M, L, XL",
                img: "{b64_04}",
                fallback_img: "Image/04 (Female).jpeg",
                measurements: "Length: XS(60cm), S(62cm), M(64cm), L(66cm), XL(68cm)\\n1/2 Chest: XS(44cm), S(46cm), M(48cm), L(50cm), XL(52cm)\\nSleeve: XS(57cm), S(58cm), M(59cm), L(60cm), XL(61cm)"
            }},
            "05": {{
                code: "05",
                name: "Women's Semi Long Cardigan Sweater",
                color: "Solid Black with Ethnic Border Trim",
                gender: "Women's",
                sizes: "S, M, L, XL, XXL",
                img: "{b64_05}",
                fallback_img: "Image/05 (Female).jpeg",
                measurements: "Length: S(64cm), M(66cm), L(68cm), XL(70cm), XXL(72cm)\\n1/2 Chest: S(51cm), M(53cm), L(55cm), XL(57cm), XXL(59cm)\\nSleeve: S(52cm), M(53cm), L(54cm), XL(55cm), XXL(56cm)"
            }}
        }};

        function getSweaterMeta(key) {{
            if (!key) return null;
            if (SWEATER_DETAILS[key]) return SWEATER_DETAILS[key];
            if (key.includes('01')) return SWEATER_DETAILS['01'];
            if (key.includes('02')) return SWEATER_DETAILS['02'];
            if (key.includes('03')) return SWEATER_DETAILS['03'];
            if (key.includes('04')) return SWEATER_DETAILS['04'];
            if (key.includes('05')) return SWEATER_DETAILS['05'];
            return null;
        }}

        let store = JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{{}}');
        let regionLocks = JSON.parse(localStorage.getItem('EXIUM_REGION_LOCKS') || '{{}}');
        let isGlobalAccessOpen = JSON.parse(localStorage.getItem('EXIUM_GLOBAL_ACCESS') || 'true');

        let currentRegionCode = null;
        let activeTerritoryIndex = 0;
        let isAdminLoggedIn = false;
        let pendingDeleteRegionCode = null;

        // Auto-Restore Session on Refresh
        window.addEventListener('DOMContentLoaded', () => {{
            populateZoneDropdown();
            checkGlobalLockBanner();

            const savedSession = JSON.parse(localStorage.getItem('EXIUM_ACTIVE_SESSION') || 'null');
            if (savedSession && savedSession.region_code && REGION_MAP[savedSession.region_code]) {{
                unlockRegion(savedSession.region_code, true);
                if (typeof savedSession.territory_idx === 'number') {{
                    selectTerritoryTab(savedSession.territory_idx, false);
                }}
            }}
        }});

        function handlePasswordKey(e) {{
            if (e.key === 'Enter') {{
                e.preventDefault();
                unlockRegion();
            }}
        }}

        function handleAdminPasswordKey(e) {{
            if (e.key === 'Enter') {{
                e.preventDefault();
                loginAdmin();
            }}
        }}

        function checkGlobalLockBanner() {{
            const banner = document.getElementById('login-global-locked-alert');
            if (!isGlobalAccessOpen) {{
                banner.classList.remove('hidden');
            }} else {{
                banner.classList.add('hidden');
            }}
        }}

        function populateZoneDropdown() {{
            const select = document.getElementById('select-zone');
            ZONES.forEach(z => {{
                const opt = document.createElement('option');
                opt.value = z;
                opt.textContent = z;
                select.appendChild(opt);
            }});
        }}

        function onZoneChanged() {{
            const zone = document.getElementById('select-zone').value;
            const regSelect = document.getElementById('select-region');
            regSelect.innerHTML = '<option value="">-- Select Region --</option>';
            
            document.getElementById('rh-info-card').classList.add('hidden');
            document.getElementById('password-section').classList.add('hidden');
            document.getElementById('unlock-btn-container').classList.add('hidden');
            document.getElementById('region-password').value = '';

            if (!zone) {{
                regSelect.disabled = true;
                return;
            }}

            const matchingRegions = Object.values(REGION_MAP).filter(r => r.zone === zone);
            matchingRegions.sort((a, b) => a.region_name.localeCompare(b.region_name));

            matchingRegions.forEach(r => {{
                const opt = document.createElement('option');
                opt.value = r.sap_region_code;
                opt.textContent = r.region_name;
                regSelect.appendChild(opt);
            }});

            regSelect.disabled = false;
        }}

        function onRegionChanged() {{
            const code = document.getElementById('select-region').value;
            if (!code) {{
                document.getElementById('rh-info-card').classList.add('hidden');
                document.getElementById('password-section').classList.add('hidden');
                document.getElementById('unlock-btn-container').classList.add('hidden');
                return;
            }}

            const r = REGION_MAP[code];
            document.getElementById('rh-name-display').textContent = r.regional_head;
            document.getElementById('rh-territory-count').textContent = `${{r.territories.length}} Territories under this Region`;
            document.getElementById('rh-info-card').classList.remove('hidden');

            document.getElementById('password-section').classList.remove('hidden');
            document.getElementById('unlock-btn-container').classList.remove('hidden');
            document.getElementById('region-password').value = '';
            document.getElementById('region-password').focus();
        }}

        function unlockRegion(bypassCode = null, isRestoringSession = false) {{
            const code = bypassCode || document.getElementById('select-region').value;
            const pass = document.getElementById('region-password').value.trim();

            if (!bypassCode && pass !== code && pass !== 'Exium MUPS' && pass !== 'admin2026') {{
                alert('Invalid Password! Please enter the correct password.');
                return;
            }}

            if (!isAdminLoggedIn && !bypassCode) {{
                if (!isGlobalAccessOpen) {{
                    alert('Submissions are currently closed by Central Admin.');
                    return;
                }}
            }}

            currentRegionCode = code;
            const r = REGION_MAP[code];

            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify({{
                region_code: code,
                territory_idx: isRestoringSession ? activeTerritoryIndex : 0
            }}));

            const adminDesk = document.getElementById('header-admin-btn-container');
            const adminMob = document.getElementById('mobile-admin-btn-container');
            if (adminDesk) adminDesk.classList.add('hidden');
            if (adminMob) adminMob.classList.add('hidden');

            document.getElementById('banner-zone').textContent = r.zone;
            document.getElementById('banner-region').textContent = r.region_name;
            document.getElementById('banner-rh').innerHTML = `<i class="fa-solid fa-user-tie text-orange-500"></i> <span class="truncate">${{r.regional_head}}</span>`;

            document.getElementById('selection-view').classList.add('hidden');
            document.getElementById('workspace-view').classList.remove('hidden');

            updateLockStatusUI();
            renderTerritoryTabs();
            if (!isRestoringSession) {{
                selectTerritoryTab(0, true);
            }}
        }}

        function exitRegionWorkspace() {{
            saveCurrentRegionData(false);
            localStorage.removeItem('EXIUM_ACTIVE_SESSION');
            currentRegionCode = null;
            document.getElementById('workspace-view').classList.add('hidden');
            document.getElementById('selection-view').classList.remove('hidden');

            const adminDesk = document.getElementById('header-admin-btn-container');
            const adminMob = document.getElementById('mobile-admin-btn-container');
            if (adminDesk) adminDesk.classList.remove('hidden');
            if (adminMob) adminMob.classList.remove('hidden');
        }}

        function isRegionLocked() {{
            if (!isGlobalAccessOpen) return true;
            return !!regionLocks[currentRegionCode];
        }}

        function updateLockStatusUI() {{
            const locked = isRegionLocked();
            const badge = document.getElementById('banner-locked-status');
            if (locked) {{
                badge.classList.remove('hidden');
            }} else {{
                badge.classList.add('hidden');
            }}
        }}

        function renderTerritoryTabs() {{
            const r = REGION_MAP[currentRegionCode];
            const container = document.getElementById('territory-tabs-container');
            const mobileSelect = document.getElementById('mobile-territory-select');
            
            container.innerHTML = '';
            mobileSelect.innerHTML = '';

            let completedCount = 0;

            r.territories.forEach((t, idx) => {{
                const d = store[t.sap_territory_code] || {{}};
                const status = getTerritoryStatus(d);
                if (status === 'Complete') completedCount++;

                const btn = document.createElement('button');
                btn.className = `w-full text-left p-2.5 rounded-xl border text-xs flex items-center justify-between transition ${{
                    idx === activeTerritoryIndex 
                        ? 'bg-slate-900 border-slate-900 text-white font-black shadow-md ring-2 ring-orange-500/50' 
                        : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                }}`;

                let badge = '<span class="text-[10px] text-slate-400 font-medium">Pending</span>';
                if (status === 'Complete') badge = `<span class="text-[10px] ${{idx === activeTerritoryIndex ? 'bg-emerald-500 text-slate-950' : 'bg-emerald-100 text-emerald-800'}} font-bold px-2 py-0.5 rounded-full"><i class="fa-solid fa-check"></i> Done</span>`;
                else if (status === 'In Progress') badge = `<span class="text-[10px] ${{idx === activeTerritoryIndex ? 'bg-amber-400 text-slate-950' : 'bg-amber-100 text-amber-800'}} font-semibold px-2 py-0.5 rounded-full">In Progress</span>`;

                btn.innerHTML = `
                    <div class="truncate pr-2">
                        <div class="truncate font-bold">${{t.territory_name}}</div>
                        <div class="text-[10px] ${{idx === activeTerritoryIndex ? 'text-slate-300' : 'text-slate-500'}} font-normal">${{t.sap_territory_code}}</div>
                    </div>
                    ${{badge}}
                `;

                btn.onclick = () => selectTerritoryTab(idx, true);
                container.appendChild(btn);

                const mOpt = document.createElement('option');
                mOpt.value = idx;
                mOpt.textContent = `${{t.territory_name}} (${{status === 'Complete' ? 'Done' : status}})`;
                if (idx === activeTerritoryIndex) mOpt.selected = true;
                mobileSelect.appendChild(mOpt);
            }});

            const badgeText = `${{completedCount}}/${{r.territories.length}} Done`;
            const badgeClass = `text-[11px] font-bold px-2.5 py-0.5 rounded-full ${{
                completedCount === r.territories.length ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' : 'bg-slate-100 text-slate-700'
            }}`;

            document.getElementById('region-progress-badge').textContent = badgeText;
            document.getElementById('region-progress-badge').className = badgeClass;
            document.getElementById('mobile-progress-badge').textContent = badgeText;
            document.getElementById('mobile-progress-badge').className = badgeClass;
        }}

        function selectTerritoryTab(idx, shouldScroll = true) {{
            activeTerritoryIndex = idx;
            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[idx];
            const d = store[t.sap_territory_code] || {{}};

            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify({{
                region_code: currentRegionCode,
                territory_idx: idx
            }}));

            document.getElementById('mobile-territory-select').value = idx;
            document.getElementById('current-territory-title').textContent = t.territory_name;
            document.getElementById('current-territory-code').textContent = `SAP Code: ${{t.sap_territory_code}}`;

            const status = getTerritoryStatus(d);
            const statusBadge = document.getElementById('current-territory-status');
            statusBadge.textContent = status;
            statusBadge.className = `text-[9px] sm:text-[10px] font-bold px-2 py-0.2 rounded-full ${{
                status === 'Complete' ? 'bg-emerald-500 text-slate-950 font-black' :
                status === 'In Progress' ? 'bg-amber-400 text-slate-950 font-bold' :
                'bg-white/10 text-slate-200 border border-white/20'
            }}`;

            const isLocked = isRegionLocked();
            const lockedNotice = document.getElementById('territory-locked-notice');
            if (isLocked) {{
                lockedNotice.classList.remove('hidden');
            }} else {{
                lockedNotice.classList.add('hidden');
            }}

            // Populate C1
            const c1DocInput = document.getElementById('c1_doc_name');
            c1DocInput.value = d.c1_doc_name || '';
            c1DocInput.disabled = isLocked;

            ['m1', 'm2', 'm3', 'm4'].forEach(m => {{
                const sw = d[`c1_${{m}}_sweater`] || '';
                const sz = d[`c1_${{m}}_size`] || '';
                const swSel = document.getElementById(`c1_${{m}}_sweater`);
                const szSel = document.getElementById(`c1_${{m}}_size`);
                
                swSel.value = sw;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c1_${{m}}_sweater`, `c1_${{m}}_size`, sz);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c1_${{m}}_img_preview`, sw);
            }});

            // Populate C2
            ['d1', 'd2', 'd3', 'd4'].forEach(d_item => {{
                const dNameInput = document.getElementById(`c2_${{d_item}}_name`);
                dNameInput.value = d[`c2_${{d_item}}_name`] || '';
                dNameInput.disabled = isLocked;

                const sw = d[`c2_${{d_item}}_sweater`] || '';
                const sz = d[`c2_${{d_item}}_size`] || '';
                const swSel = document.getElementById(`c2_${{d_item}}_sweater`);
                const szSel = document.getElementById(`c2_${{d_item}}_size`);

                swSel.value = sw;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c2_${{d_item}}_sweater`, `c2_${{d_item}}_size`, sz);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c2_${{d_item}}_img_preview`, sw);
            }});

            renderTerritoryTabs();

            // Auto-scroll to top of active territory card
            if (shouldScroll) {{
                window.scrollTo({{ top: 0, behavior: 'smooth' }});
            }}
        }}

        function onSweaterSelectChange(slotId, val) {{
            updateSizeOptionsForSelect(`${{slotId}}_sweater`, `${{slotId}}_size`);
            updateSlotImagePreview(`${{slotId}}_img_preview`, val);
            onDataChanged();
        }}

        function updateSizeOptionsForSelect(sweaterSelectId, sizeSelectId, preselectedVal = '') {{
            const swVal = document.getElementById(sweaterSelectId).value;
            const sizeSelect = document.getElementById(sizeSelectId);
            
            let sizes = ['S', 'M', 'L', 'XL', 'XXL'];
            if (swVal.includes('04')) {{
                sizes = ['XS', 'S', 'M', 'L', 'XL'];
            }}

            sizeSelect.innerHTML = '<option value="">-- Size --</option>';
            sizes.forEach(sz => {{
                const opt = document.createElement('option');
                opt.value = sz;
                opt.textContent = sz;
                if (sz === preselectedVal) opt.selected = true;
                sizeSelect.appendChild(opt);
            }});
        }}

        function updateSlotImagePreview(previewContainerId, swVal) {{
            const container = document.getElementById(previewContainerId);
            const item = getSweaterMeta(swVal);
            if (item) {{
                container.innerHTML = `
                    <img src="${{item.img}}" onerror="this.src='${{item.fallback_img}}'" class="w-full h-full object-cover rounded-xl" style="image-rendering: -webkit-optimize-contrast;" alt="${{item.name}}">
                    <div class="absolute inset-0 bg-slate-900/30 opacity-0 group-hover:opacity-100 flex items-center justify-center transition rounded-xl">
                        <i class="fa-solid fa-magnifying-glass-plus text-white text-base"></i>
                    </div>
                `;
            }} else {{
                container.innerHTML = '<i class="fa-solid fa-shirt text-lg text-slate-300"></i>';
            }}
        }}

        function zoomSlotImage(selectId) {{
            const val = document.getElementById(selectId).value;
            if (val) {{
                openImageLightbox(val);
            }} else {{
                openCatalogModal();
            }}
        }}

        function onDataChanged() {{
            if (isRegionLocked()) return;

            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[activeTerritoryIndex];
            const terrCode = t.sap_territory_code;

            store[terrCode] = {{
                c1_doc_name: document.getElementById('c1_doc_name').value.trim(),
                c1_m1_sweater: document.getElementById('c1_m1_sweater').value,
                c1_m1_size: document.getElementById('c1_m1_size').value,
                c1_m2_sweater: document.getElementById('c1_m2_sweater').value,
                c1_m2_size: document.getElementById('c1_m2_size').value,
                c1_m3_sweater: document.getElementById('c1_m3_sweater').value,
                c1_m3_size: document.getElementById('c1_m3_size').value,
                c1_m4_sweater: document.getElementById('c1_m4_sweater').value,
                c1_m4_size: document.getElementById('c1_m4_size').value,

                c2_d1_name: document.getElementById('c2_d1_name').value.trim(),
                c2_d1_sweater: document.getElementById('c2_d1_sweater').value,
                c2_d1_size: document.getElementById('c2_d1_size').value,
                c2_d2_name: document.getElementById('c2_d2_name').value.trim(),
                c2_d2_sweater: document.getElementById('c2_d2_sweater').value,
                c2_d2_size: document.getElementById('c2_d2_size').value,
                c2_d3_name: document.getElementById('c2_d3_name').value.trim(),
                c2_d3_sweater: document.getElementById('c2_d3_sweater').value,
                c2_d3_size: document.getElementById('c2_d3_size').value,
                c2_d4_name: document.getElementById('c2_d4_name').value.trim(),
                c2_d4_sweater: document.getElementById('c2_d4_sweater').value,
                c2_d4_size: document.getElementById('c2_d4_size').value,
            }};

            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));
            renderTerritoryTabs();
        }}

        function saveCurrentTerritoryClick() {{
            onDataChanged();
            alert('✅ Territory information saved successfully!');
        }}

        function getTerritoryStatus(d) {{
            if (!d) return 'Not Started';
            const c1Ok = d.c1_doc_name && d.c1_m1_sweater && d.c1_m1_size && d.c1_m2_sweater && d.c1_m2_size && d.c1_m3_sweater && d.c1_m3_size && d.c1_m4_sweater && d.c1_m4_size;
            const c2Ok = d.c2_d1_name && d.c2_d1_sweater && d.c2_d1_size && d.c2_d2_name && d.c2_d2_sweater && d.c2_d2_size && d.c2_d3_name && d.c2_d3_sweater && d.c2_d3_size && d.c2_d4_name && d.c2_d4_sweater && d.c2_d4_size;

            if (c1Ok && c2Ok) return 'Complete';
            if (d.c1_doc_name || d.c2_d1_name || d.c2_d2_name || d.c2_d3_name || d.c2_d4_name) return 'In Progress';
            return 'Not Started';
        }}

        function navigateTerritory(dir) {{
            onDataChanged();
            const r = REGION_MAP[currentRegionCode];
            const nextIdx = activeTerritoryIndex + dir;
            if (nextIdx >= 0 && nextIdx < r.territories.length) {{
                selectTerritoryTab(nextIdx, true);
            }}
        }}

        function saveCurrentRegionData(showAlert = true) {{
            onDataChanged();
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));
            if (showAlert) {{
                alert('✅ Information successfully saved!');
            }}
        }}

        // Full-Resolution Lightbox Zoom Modal (100% Reliable)
        function openImageLightbox(key) {{
            const item = getSweaterMeta(key);
            if (!item) return;

            const modal = document.getElementById('image-lightbox-modal');
            const imgEl = document.getElementById('lightbox-img');
            imgEl.src = item.img;
            imgEl.onerror = function() {{ this.src = item.fallback_img; }};
            
            document.getElementById('lightbox-code-badge').textContent = "Code: " + item.code;
            document.getElementById('lightbox-gender').textContent = item.gender;
            document.getElementById('lightbox-title').textContent = item.name;
            document.getElementById('lightbox-color').textContent = item.color;
            document.getElementById('lightbox-sizes').textContent = item.sizes;
            document.getElementById('lightbox-measurements').innerText = item.measurements;

            modal.classList.remove('hidden');
        }}

        function closeImageLightbox() {{
            document.getElementById('image-lightbox-modal').classList.add('hidden');
        }}

        // Catalogue Modal
        function openCatalogModal() {{ document.getElementById('catalog-modal').classList.remove('hidden'); }}
        function closeCatalogModal() {{ document.getElementById('catalog-modal').classList.add('hidden'); }}

        // Admin Modal
        function openAdminModal() {{
            document.getElementById('admin-modal').classList.remove('hidden');
            if (isAdminLoggedIn) {{
                renderAdminDashboard();
            }} else {{
                document.getElementById('admin-login-box').classList.remove('hidden');
                document.getElementById('admin-dashboard-content').classList.add('hidden');
                document.getElementById('admin-pass-input').value = '';
                document.getElementById('admin-pass-input').focus();
            }}
        }}
        function closeAdminModal() {{
            isAdminLoggedIn = false;
            document.getElementById('admin-modal').classList.add('hidden');
        }}

        // Admin Login
        function loginAdmin() {{
            const pass = document.getElementById('admin-pass-input').value.trim();
            if (pass === 'Exium MUPS' || pass.toLowerCase() === 'exium mups' || pass === 'admin2026') {{
                isAdminLoggedIn = true;
                document.getElementById('admin-login-box').classList.add('hidden');
                document.getElementById('admin-dashboard-content').classList.remove('hidden');
                renderAdminDashboard();
            }} else {{
                alert('Invalid Admin Password!');
            }}
        }}

        function toggleGlobalAccess() {{
            isGlobalAccessOpen = !isGlobalAccessOpen;
            localStorage.setItem('EXIUM_GLOBAL_ACCESS', JSON.stringify(isGlobalAccessOpen));
            
            const badge = document.getElementById('admin-global-access-badge');
            const btn = document.getElementById('admin-global-toggle-btn');

            if (isGlobalAccessOpen) {{
                badge.textContent = 'OPEN';
                badge.className = 'px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-200';
                btn.className = 'px-3 sm:px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white font-bold rounded-xl text-xs transition flex items-center gap-1.5 shadow-sm active:scale-95';
                btn.innerHTML = '<i class="fa-solid fa-lock"></i><span>Lock Submissions</span>';
            }} else {{
                badge.textContent = 'LOCKED';
                badge.className = 'px-2.5 py-0.5 rounded-full text-xs font-bold bg-rose-100 text-rose-800 border border-rose-200';
                btn.className = 'px-3 sm:px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold rounded-xl text-xs transition flex items-center gap-1.5 shadow-sm active:scale-95';
                btn.innerHTML = '<i class="fa-solid fa-lock-open"></i><span>Re-Open Submissions</span>';
            }}

            checkGlobalLockBanner();
            alert(isGlobalAccessOpen ? 'Submissions are now OPEN for all regions.' : 'Submissions are now LOCKED globally.');
        }}

        function toggleSpecificRegionLock(code) {{
            regionLocks[code] = !regionLocks[code];
            localStorage.setItem('EXIUM_REGION_LOCKS', JSON.stringify(regionLocks));
            renderAdminRegionTable();
        }}

        // Admin Delete Single Region Data
        function promptDeleteRegionData(code) {{
            pendingDeleteRegionCode = code;
            const r = REGION_MAP[code];
            document.getElementById('delete-confirm-message').innerHTML = `Are you sure you want to delete all saved submission data for:<br><strong class="text-slate-900">${{r.region_name}} (${{r.sap_region_code}})</strong>?<br><span class="text-[11px] text-rose-600 font-semibold">This action cannot be undone.</span>`;
            document.getElementById('delete-confirm-btn').onclick = executeDeleteRegionData;
            document.getElementById('delete-confirm-modal').classList.remove('hidden');
        }}

        function executeDeleteRegionData() {{
            if (!pendingDeleteRegionCode) return;
            const r = REGION_MAP[pendingDeleteRegionCode];
            r.territories.forEach(t => {{
                delete store[t.sap_territory_code];
            }});
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));
            closeDeleteConfirmModal();
            renderAdminDashboard();
            alert(`✅ All data for region ${{r.region_name}} has been cleared.`);
        }}

        // Admin Delete ALL Submissions Data
        function promptDeleteAllData() {{
            document.getElementById('delete-confirm-message').innerHTML = `<strong class="text-rose-600 text-sm">⚠️ DANGER: DELETE ALL DATA</strong><br><br>Are you sure you want to delete <strong class="text-slate-900">ALL submission data</strong> across all 252 regions and 1,856 territories?<br><span class="text-[11px] text-rose-600 font-semibold">All entered doctor names, sweater selections, and sizes will be permanently wiped out!</span>`;
            document.getElementById('delete-confirm-btn').onclick = executeDeleteAllData;
            document.getElementById('delete-confirm-modal').classList.remove('hidden');
        }}

        function executeDeleteAllData() {{
            store = {{}};
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));
            closeDeleteConfirmModal();
            renderAdminDashboard();
            alert('✅ All campaign submission data has been completely wiped out.');
        }}

        function closeDeleteConfirmModal() {{
            pendingDeleteRegionCode = null;
            document.getElementById('delete-confirm-modal').classList.add('hidden');
        }}

        function renderAdminDashboard() {{
            let completedTerritories = 0;
            let inProgressTerritories = 0;
            let totalSweatersCount = 0;

            const counts = {{
                "01 - Men's V-Neck (Grey)": {{ XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }},
                "02 - Men's V-Neck (Navy Blue)": {{ XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }},
                "03 - Men's V-Neck (Cream Check)": {{ XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }},
                "04 - Women's Short Cardigan (Check)": {{ XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }},
                "05 - Women's Semi Long Cardigan (Black)": {{ XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }}
            }};

            ALL_TERRITORIES.forEach(t => {{
                const d = store[t['SAP Territory Code']];
                if (d) {{
                    const st = getTerritoryStatus(d);
                    if (st === 'Complete') completedTerritories++;
                    else if (st === 'In Progress') inProgressTerritories++;

                    ['m1', 'm2', 'm3', 'm4'].forEach(m => {{
                        const sw = d[`c1_${{m}}_sweater`];
                        const sz = d[`c1_${{m}}_size`];
                        if (sw && sz && counts[sw]) {{
                            counts[sw][sz] = (counts[sw][sz] || 0) + 1;
                            counts[sw].total++;
                            totalSweatersCount++;
                        }}
                    }});

                    ['d1', 'd2', 'd3', 'd4'].forEach(d_item => {{
                        const sw = d[`c2_${{d_item}}_sweater`];
                        const sz = d[`c2_${{d_item}}_size`];
                        if (sw && sz && counts[sw]) {{
                            counts[sw][sz] = (counts[sw][sz] || 0) + 1;
                            counts[sw].total++;
                            totalSweatersCount++;
                        }}
                    }});
                }}
            }});

            const pct = Math.round((completedTerritories / ALL_TERRITORIES.length) * 100);
            document.getElementById('admin-kpi-completed').textContent = completedTerritories;
            document.getElementById('admin-kpi-completed-pct').textContent = `${{pct}}% of total`;
            document.getElementById('admin-kpi-inprogress').textContent = inProgressTerritories;
            document.getElementById('admin-kpi-sweaters').textContent = `${{totalSweatersCount.toLocaleString()}} / 14,848`;

            let matrixHtml = `
                <table class="w-full text-xs text-center border-collapse border border-slate-200 rounded-xl overflow-hidden shadow-sm min-w-[500px]">
                    <thead>
                        <tr class="bg-slate-900 text-white">
                            <th class="py-2 px-3 text-left">Sweater Design</th>
                            <th>XS</th><th>S</th><th>M</th><th>L</th><th>XL</th><th>XXL</th>
                            <th class="bg-orange-600 text-white font-bold">Total Pieces</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-200 bg-white text-slate-700">
            `;

            let szTotals = {{ XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0 }};
            Object.keys(counts).forEach(k => {{
                const row = counts[k];
                Object.keys(szTotals).forEach(sz => szTotals[sz] += row[sz]);
                matrixHtml += `
                    <tr class="hover:bg-slate-50">
                        <td class="py-1.5 px-3 text-left font-semibold text-slate-900">${{k}}</td>
                        <td>${{row.XS || '-'}}</td><td>${{row.S || '-'}}</td><td>${{row.M || '-'}}</td>
                        <td>${{row.L || '-'}}</td><td>${{row.XL || '-'}}</td><td>${{row.XXL || '-'}}</td>
                        <td class="font-bold text-orange-600 bg-orange-50 font-mono">${{row.total}}</td>
                    </tr>
                `;
            }});

            matrixHtml += `
                    </tbody>
                    <tfoot>
                        <tr class="bg-slate-900 text-white font-bold">
                            <td class="py-2 px-3 text-left">GRAND TOTAL</td>
                            <td>${{szTotals.XS}}</td><td>${{szTotals.S}}</td><td>${{szTotals.M}}</td>
                            <td>${{szTotals.L}}</td><td>${{szTotals.XL}}</td><td>${{szTotals.XXL}}</td>
                            <td class="bg-emerald-600 text-white text-sm font-mono">${{totalSweatersCount.toLocaleString()}}</td>
                        </tr>
                    </tfoot>
                </table>
            `;

            document.getElementById('admin-procurement-matrix-container').innerHTML = matrixHtml;
            renderAdminRegionTable();
        }}

        function renderAdminRegionTable() {{
            const tbody = document.getElementById('admin-region-tbody');
            const q = (document.getElementById('admin-region-search').value || '').toLowerCase().trim();
            tbody.innerHTML = '';

            Object.values(REGION_MAP).forEach(r => {{
                if (q && !r.region_name.toLowerCase().includes(q) && !r.regional_head.toLowerCase().includes(q) && !r.sap_region_code.includes(q)) {{
                    return;
                }}

                let compCount = 0;
                r.territories.forEach(t => {{
                    if (getTerritoryStatus(store[t.sap_territory_code]) === 'Complete') compCount++;
                }});

                const isLocked = !isGlobalAccessOpen || !!regionLocks[r.sap_region_code];

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="py-2 px-3 font-mono text-slate-600">${{r.sap_region_code}}</td>
                    <td class="py-2 px-3 font-semibold text-slate-900">${{r.region_name}}</td>
                    <td class="py-2 px-3 text-orange-700 font-medium">${{r.regional_head}}</td>
                    <td class="py-2 px-2 text-center">
                        <span class="px-2 py-0.5 rounded-full text-[10px] ${{
                            compCount === r.territories.length ? 'bg-emerald-100 text-emerald-800 font-bold' :
                            compCount > 0 ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-600'
                        }}">${{compCount}}/${{r.territories.length}}</span>
                    </td>
                    <td class="py-2 px-3 text-right space-x-1 whitespace-nowrap">
                        <button onclick="toggleSpecificRegionLock('${{r.sap_region_code}}')" title="Lock/Unlock Region" class="px-2 py-1 rounded text-[10px] font-bold ${{
                            isLocked ? 'bg-rose-100 text-rose-800 border border-rose-200' : 'bg-slate-100 text-slate-700 hover:bg-slate-200 border border-slate-200'
                        }}">
                            ${{isLocked ? '<i class="fa-solid fa-lock"></i>' : '<i class="fa-solid fa-lock-open"></i>'}}
                        </button>
                        <button onclick="promptDeleteRegionData('${{r.sap_region_code}}')" title="Delete Region Data" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded text-[10px] font-bold">
                            <i class="fa-solid fa-trash-can"></i>
                        </button>
                        <button onclick="adminDirectUnlock('${{r.sap_region_code}}')" title="Open Workspace" class="px-2.5 py-1 bg-slate-900 hover:bg-slate-800 text-white rounded text-[10px] font-bold">
                            Open
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            }});
        }}

        function adminDirectUnlock(regionCode) {{
            closeAdminModal();
            unlockRegion(regionCode, false);
        }}

        // Master Excel Export
        function exportMasterExcelFromAdmin() {{
            exportMasterExcelFile("Exium_MUPS_Sweater_Campaign_2026_Master_Export.xlsx");
        }}

        function exportMasterExcelFile(filename) {{
            const c1Rows = [];
            const c2Rows = [];

            ALL_TERRITORIES.forEach(t => {{
                const terrCode = t['SAP Territory Code'];
                const d = store[terrCode] || {{}};

                c1Rows.push({{
                    "Zone": t.Zone,
                    "SAP Region Code": t['SAP Region Code'],
                    "Region": t.Region,
                    "Regional Head": t['Regional Head'],
                    "SAP Territory Code": t['SAP Territory Code'],
                    "Territory": t.Territory,
                    "Doctor Name": d.c1_doc_name || '',
                    "Sweater 1": d.c1_m1_sweater || '',
                    "Size 1": d.c1_m1_size || '',
                    "Sweater 2": d.c1_m2_sweater || '',
                    "Size 2": d.c1_m2_size || '',
                    "Sweater 3": d.c1_m3_sweater || '',
                    "Size 3": d.c1_m3_size || '',
                    "Sweater 4": d.c1_m4_sweater || '',
                    "Size 4": d.c1_m4_size || '',
                    "Status": (d.c1_doc_name && d.c1_m1_sweater && d.c1_m1_size && d.c1_m2_sweater && d.c1_m2_size && d.c1_m3_sweater && d.c1_m3_size && d.c1_m4_sweater && d.c1_m4_size) ? "Complete" : (d.c1_doc_name ? "In Progress" : "Not Started")
                }});

                c2Rows.push({{
                    "Zone": t.Zone,
                    "SAP Region Code": t['SAP Region Code'],
                    "Region": t.Region,
                    "Regional Head": t['Regional Head'],
                    "SAP Territory Code": t['SAP Territory Code'],
                    "Territory": t.Territory,
                    "Doctor 1 Name": d.c2_d1_name || '',
                    "Sweater 1": d.c2_d1_sweater || '',
                    "Size 1": d.c2_d1_size || '',
                    "Doctor 2 Name": d.c2_d2_name || '',
                    "Sweater 2": d.c2_d2_sweater || '',
                    "Size 2": d.c2_d2_size || '',
                    "Doctor 3 Name": d.c2_d3_name || '',
                    "Sweater 3": d.c2_d3_sweater || '',
                    "Size 3": d.c2_d3_size || '',
                    "Doctor 4 Name": d.c2_d4_name || '',
                    "Sweater 4": d.c2_d4_sweater || '',
                    "Size 4": d.c2_d4_size || '',
                    "Status": (d.c2_d1_name && d.c2_d1_sweater && d.c2_d1_size && d.c2_d2_name && d.c2_d2_sweater && d.c2_d2_size && d.c2_d3_name && d.c2_d3_sweater && d.c2_d3_size && d.c2_d4_name && d.c2_d4_sweater && d.c2_d4_size) ? "Complete" : (d.c2_d1_name ? "In Progress" : "Not Started")
                }});
            }});

            const wb = XLSX.utils.book_new();
            const ws1 = XLSX.utils.json_to_sheet(c1Rows);
            const ws2 = XLSX.utils.json_to_sheet(c2Rows);

            XLSX.utils.book_append_sheet(wb, ws1, "Gyne Core Doctor (Family)");
            XLSX.utils.book_append_sheet(wb, ws2, "Core Doctor Maximization");

            XLSX.writeFile(wb, filename);
            alert(`📥 Excel file "${{filename}}" downloaded with all current saved inputs!`);
        }}
    </script>
</body>
</html>
'''

with open('Sweater_Campaign_Portal.html', 'w', encoding='utf-8') as f:
    f.write(html_code)

print('Updated Sweater_Campaign_Portal.html with 1-line mobile logo+title, row 2 Catalogue button, and Sticky/Frozen Active Territory Bar generated successfully!')
