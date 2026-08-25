import os
import json
import base64
import pandas as pd

excel_file = r"G:\Exium\2026\4Q'26\Sweater\FF list.xlsx"
df_ff = pd.read_excel(excel_file)
all_territories = df_ff.to_dict(orient='records')

zones = sorted(list(set(df_ff['Zone'].dropna().astype(str))))
region_map = {}

for _, row in df_ff.iterrows():
    reg_code = str(row['SAP Region Code']).strip()
    if reg_code not in region_map:
        region_map[reg_code] = {
            'sap_region_code': reg_code,
            'region_name': str(row['Region']).strip(),
            'regional_head': str(row['Regional Head']).strip(),
            'zone': str(row['Zone']).strip(),
            'territories': []
        }
    region_map[reg_code]['territories'].append({
        'sap_territory_code': str(row['SAP Territory Code']).strip(),
        'territory_name': str(row['Territory']).strip(),
        'zone': str(row['Zone']).strip(),
        'region': str(row['Region']).strip(),
        'regional_head': str(row['Regional Head']).strip()
    })

img_dir = r"G:\Exium\2026\4Q'26\Sweater\Image"
images_b64 = {}
img_map = {
    "01": "01 (Men).jpeg",
    "02": "02 (Men).jpeg",
    "03": "03 (Men).jpeg",
    "04": "04 (Female).jpeg",
    "05": "05 (Female).jpeg"
}

print("Encoding images...")
for k, fname in img_map.items():
    fpath = os.path.join(img_dir, fname)
    if os.path.exists(fpath):
        with open(fpath, "rb") as img_f:
            images_b64[k] = f"data:image/jpeg;base64,{base64.b64encode(img_f.read()).decode('utf-8')}"
    else:
        images_b64[k] = f"Image/{fname}"

zone_options_html = '<option value="">-- Select Zone --</option>\n'
for z in zones:
    zone_options_html += f'                            <option value="{z}">{z}</option>\n'

territories_json = json.dumps(all_territories)
region_map_json = json.dumps(region_map)
zones_json = json.dumps(zones)

html_template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Exium MUPS - 4Q'26 Doctor Sweater Campaign Portal</title>
    
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#fff7ed',
                            100: '#ffedd5',
                            500: '#f97316',
                            600: '#ea580c',
                            700: '#c2410c',
                            800: '#9a3412',
                            900: '#7c2d12',
                        },
                        exium: {
                            dark: '#0f172a',
                            card: '#1e293b',
                            teal: '#0d9488',
                            purple: '#7e22ce'
                        }
                    }
                }
            }
        }
    </script>
    
    <!-- FontAwesome & SheetJS (xlsx) -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    
    <!-- Inter Font -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #0b0f19;
            color: #f1f5f9;
            -webkit-tap-highlight-color: transparent;
        }
        .custom-scrollbar::-webkit-scrollbar {
            width: 5px;
            height: 5px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
            background: rgba(15, 23, 42, 0.6);
            border-radius: 9999px;
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
            background: #f97316;
            border-radius: 9999px;
        }
        .sweater-card-img {
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .sweater-card-img:hover {
            transform: scale(1.03);
            box-shadow: 0 8px 20px -4px rgba(249, 115, 22, 0.3);
        }
        input:focus, select:focus {
            box-shadow: 0 0 0 2px rgba(249, 115, 22, 0.4);
        }
        .nav-tab-active {
            background-color: #f97316 !important;
            color: #ffffff !important;
            font-weight: 700 !important;
            border-color: #f97316 !important;
        }
    </style>
</head>
<body class="min-h-screen flex flex-col bg-slate-950 text-slate-100">

    <!-- GLOBAL NAVBAR -->
    <header class="sticky top-0 z-40 bg-slate-900/90 backdrop-blur-md border-b border-slate-800 shadow-lg">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
            <div class="flex items-center gap-3">
                <div class="w-10 h-10 rounded-2xl bg-gradient-to-tr from-orange-600 to-amber-500 flex items-center justify-center shadow-lg shadow-orange-500/20 text-white font-black text-lg tracking-wider">
                    EX
                </div>
                <div>
                    <h1 class="text-sm sm:text-base font-black tracking-tight text-white flex items-center gap-1.5">
                        <span>EXIUM MUPS</span>
                        <span class="text-[10px] uppercase px-2 py-0.5 rounded-full bg-orange-500/20 text-orange-400 font-bold border border-orange-500/30">4Q'26</span>
                    </h1>
                    <p class="text-[11px] sm:text-xs text-slate-400 font-medium">Doctor Sweater Gift Campaign Portal</p>
                </div>
            </div>

            <div class="flex items-center gap-2 sm:gap-3">
                <button onclick="openCatalogModal()" class="px-2.5 sm:px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold border border-slate-700 flex items-center gap-1.5 transition active:scale-95 shadow-sm">
                    <i class="fa-solid fa-shirt text-orange-400"></i>
                    <span class="hidden md:inline">Catalogue &</span> Sizes
                </button>
                <button onclick="openAdminModal()" class="px-2.5 sm:px-3 py-1.5 rounded-xl bg-gradient-to-r from-slate-800 to-slate-700 hover:from-slate-700 hover:to-slate-600 text-orange-400 text-xs font-bold border border-orange-500/30 flex items-center gap-1.5 transition active:scale-95 shadow-sm">
                    <i class="fa-solid fa-shield-halved"></i>
                    <span>Admin</span>
                </button>
            </div>
        </div>
    </header>

    <!-- MAIN BODY CONTENT -->
    <main class="flex-1 max-w-7xl w-full mx-auto p-3 sm:p-6 lg:p-8 space-y-6">

        <!-- ============================================== -->
        <!-- VIEW 1: REGIONAL MANAGER LOGIN / ZONE SELECT   -->
        <!-- ============================================== -->
        <section id="view-login" class="flex flex-col items-center justify-center py-6 sm:py-12">
            
            <div id="login-global-locked-alert" class="hidden mb-6 max-w-lg w-full bg-amber-500/10 border border-amber-500/30 rounded-2xl p-4 text-amber-400 text-xs flex items-center gap-3">
                <i class="fa-solid fa-lock text-lg"></i>
                <div>
                    <strong class="font-bold">Campaign Window Closed:</strong>
                    <span> Submissions are currently locked by the Central Admin. Viewing in Read-Only mode.</span>
                </div>
            </div>

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

                    <div id="unlock-btn-container" class="hidden pt-2">
                        <button onclick="unlockRegion()" class="w-full py-3 bg-gradient-to-r from-orange-500 to-amber-500 hover:from-orange-600 hover:to-amber-600 text-white font-bold rounded-xl text-sm shadow-lg shadow-orange-500/20 transition active:scale-[0.98] flex items-center justify-center gap-2">
                            <i class="fa-solid fa-right-to-bracket"></i>
                            <span>Unlock Region Workspace</span>
                        </button>
                    </div>
                </div>

                <div class="mt-6 pt-4 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-400">
                    <span>Default Region PIN: <code class="font-mono font-bold text-slate-600">1234</code></span>
                    <span>Admin PIN: <code class="font-mono font-bold text-slate-600">admin2026</code></span>
                </div>
            </div>
        </section>

        <!-- ============================================== -->
        <!-- VIEW 2: REGIONAL MANAGER WORKSPACE             -->
        <!-- ============================================== -->
        <section id="view-workspace" class="hidden space-y-6">

            <!-- Region Header Banner -->
            <div class="bg-white border border-slate-200 rounded-3xl p-4 sm:p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div class="space-y-1">
                    <div class="flex items-center gap-2">
                        <span id="banner-zone" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-orange-50 text-orange-700 border border-orange-200">Zone Name</span>
                        <span id="banner-region" class="px-2.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-200">SAP: 00000</span>
                    </div>
                    <h2 id="banner-rh" class="text-lg sm:text-2xl font-black text-slate-900 tracking-tight">Region: Region Name (Regional Head Name)</h2>
                    <p class="text-xs text-slate-500 flex items-center gap-2">
                        <span>Territories: <strong id="banner-total-count" class="text-slate-800">0</strong></span>
                        <span>•</span>
                        <span>Completed: <strong id="banner-complete-count" class="text-emerald-600 font-black">0</strong></span>
                    </p>
                </div>

                <div class="flex flex-wrap items-center gap-2">
                    <button onclick="exportSingleRegionExcel()" class="px-3 sm:px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold border border-slate-300 flex items-center gap-1.5 transition active:scale-95">
                        <i class="fa-solid fa-file-excel text-emerald-600 text-sm"></i>
                        <span>Export Region Excel</span>
                    </button>
                    <button onclick="exitRegionWorkspace()" class="px-3 sm:px-4 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-xl text-xs font-bold border border-rose-200 flex items-center gap-1.5 transition active:scale-95">
                        <i class="fa-solid fa-arrow-right-from-bracket"></i>
                        <span>Switch Region</span>
                    </button>
                </div>
            </div>

            <!-- Workspace Layout (Sidebar / Tabs + Territory Form) -->
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">

                <!-- Left Column: Territory Navigation -->
                <div class="lg:col-span-4 xl:col-span-3 space-y-4">
                    <div class="bg-white border border-slate-200 rounded-3xl p-4 shadow-sm space-y-3">
                        <div class="flex items-center justify-between">
                            <h3 class="text-xs font-bold uppercase tracking-wider text-slate-700 flex items-center gap-1.5">
                                <i class="fa-solid fa-map-pin text-orange-500"></i>
                                <span>Territories</span>
                            </h3>
                            <span id="region-progress-badge" class="px-2 py-0.5 rounded-full text-[10px] font-black bg-slate-100 text-slate-700">0/0</span>
                        </div>

                        <!-- Mobile Selector -->
                        <div class="block lg:hidden">
                            <select id="mobile-territory-select" onchange="selectTerritoryTab(parseInt(this.value))" class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3 py-2 text-xs font-bold text-slate-800 focus:outline-none">
                            </select>
                        </div>

                        <!-- Desktop Scrollable List -->
                        <div id="desktop-territory-list" class="hidden lg:flex flex-col gap-1.5 max-h-[620px] overflow-y-auto custom-scrollbar pr-1">
                            <!-- Populated dynamically -->
                        </div>
                    </div>
                </div>

                <!-- Right Column: Territory Data Form -->
                <div class="lg:col-span-8 xl:col-span-9 space-y-5">

                    <!-- Active Territory Header Card -->
                    <div class="bg-white border border-slate-200 rounded-3xl p-4 sm:p-5 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div class="space-y-0.5">
                            <span id="active-terr-sap" class="text-[10px] font-bold text-slate-400 uppercase tracking-widest">SAP: 00000</span>
                            <h3 id="active-terr-name" class="text-base sm:text-xl font-black text-slate-900 tracking-tight">Territory Name</h3>
                        </div>
                        <div class="flex items-center gap-2">
                            <span id="active-terr-status" class="px-3 py-1 rounded-full text-xs font-black bg-slate-100 text-slate-600 border border-slate-200 flex items-center gap-1.5">
                                <i class="fa-regular fa-circle text-[9px]"></i>
                                <span>Not Started</span>
                            </span>
                        </div>
                    </div>

                    <!-- Lock Notice if Locked -->
                    <div id="territory-locked-notice" class="hidden bg-amber-50 border border-amber-200 rounded-2xl p-3.5 text-amber-800 text-xs flex items-center gap-2.5">
                        <i class="fa-solid fa-lock text-amber-600"></i>
                        <span>This region has been locked by Admin. Changes cannot be made.</span>
                    </div>

                    <!-- ============================================== -->
                    <!-- CAMPAIGN 1: GYNE CORE DOCTOR (FAMILY PACKAGE)  -->
                    <!-- ============================================== -->
                    <div class="bg-white border-2 border-teal-500/60 rounded-3xl shadow-sm overflow-hidden">
                        <div class="bg-gradient-to-r from-teal-700 via-teal-800 to-emerald-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-8 h-8 rounded-xl bg-white/10 flex items-center justify-center font-black text-sm flex-shrink-0">1</div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Gyne Core Doctor Development (Family Package)</h4>
                                    <p class="text-[10px] sm:text-xs text-teal-100/80">1 Doctor • 3 Sweaters standard (+ Optional 4th Sweater)</p>
                                </div>
                            </div>
                            <span id="c1-status-pill" class="px-2.5 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-white/20 text-white backdrop-blur-sm self-start sm:self-auto flex-shrink-0">3 or 4 Sweaters / Terr</span>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
                            <!-- Doctor Info -->
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-3 sm:gap-4 bg-teal-50/50 p-3.5 sm:p-4 rounded-2xl border border-teal-200/80">
                                <div>
                                    <label class="block text-xs font-bold text-teal-950 uppercase tracking-wider mb-1">
                                        Doctor Name <span class="text-rose-500">*</span>
                                    </label>
                                    <input type="text" id="c1_doc_name" oninput="onDataChanged()" placeholder="Enter Gynecologist / Doctor Name..." class="w-full bg-white border border-teal-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:border-teal-500 font-medium">
                                </div>
                                <div>
                                    <div class="flex items-center justify-between mb-1">
                                        <label class="block text-xs font-bold text-teal-950 uppercase tracking-wider">
                                            Doctor RPL ID <span class="text-rose-500">*</span>
                                        </label>
                                        <span id="c1_doc_rpl_badge" class="text-[10px] font-bold text-slate-400">6 digits</span>
                                    </div>
                                    <input type="text" inputmode="numeric" maxlength="6" id="c1_doc_rpl" oninput="onRplInput(this, 'c1_doc_rpl_badge')" placeholder="6-digit RPL ID (e.g. 123456)..." class="w-full bg-white border border-teal-300 rounded-xl px-3.5 py-2 text-xs sm:text-sm text-slate-900 font-mono font-bold placeholder-slate-400 focus:outline-none focus:border-teal-500 tracking-wider">
                                </div>
                            </div>

                            <!-- 3 + Optional 4 Sweaters Grid -->
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-3 sm:gap-4">
                                <!-- Sweater 1 -->
                                <div class="bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">1</span> Doctor's Own Sweater</span>
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
                                <!-- Sweater 2 -->
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
                                <!-- Sweater 3 -->
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

                                <!-- Sweater 4: Add Button (Hidden when Card 4 is active) -->
                                <div id="c1_m4_add_container" onclick="showC1Sweater4(true)" class="border-2 border-dashed border-teal-300 hover:border-teal-500 rounded-2xl p-4 flex flex-col items-center justify-center text-center bg-teal-50/40 hover:bg-teal-50 transition cursor-pointer group min-h-[130px]">
                                    <div class="w-9 h-9 rounded-full bg-teal-100 text-teal-700 flex items-center justify-center text-sm font-black mb-1.5 group-hover:scale-110 transition shadow-sm">
                                        <i class="fa-solid fa-plus"></i>
                                    </div>
                                    <span class="text-xs font-black text-teal-900">+ Add Another Sweater</span>
                                    <span class="text-[10px] text-teal-600 font-medium">Optional 4th Sweater (Family Member)</span>
                                </div>

                                <!-- Sweater 4 Card (Shown when Add Button is clicked or data exists) -->
                                <div id="c1_m4_card" class="hidden bg-slate-50 border border-slate-200 rounded-2xl p-3 space-y-2 relative">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-teal-900 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-teal-600 text-white flex items-center justify-center text-[10px] font-black">4</span> Sweater 4 (Family Member)</span>
                                        <div class="flex items-center gap-2">
                                            <span id="c1_m4_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                            <button type="button" onclick="hideAndClearC1Sweater4()" class="text-rose-500 hover:text-rose-700 text-[11px] font-bold px-1.5 py-0.5 rounded hover:bg-rose-50 transition" title="Remove Sweater 4">
                                                <i class="fa-solid fa-xmark"></i> Remove
                                            </button>
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

                    <!-- ============================================== -->
                    <!-- CAMPAIGN 2: CORE DOCTOR MAXIMIZATION (3 DOCS)  -->
                    <!-- ============================================== -->
                    <div class="bg-white border-2 border-purple-500/60 rounded-3xl shadow-sm overflow-hidden">
                        <div class="bg-gradient-to-r from-purple-700 via-purple-800 to-indigo-800 text-white px-3.5 sm:px-6 py-2.5 sm:py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <div class="flex items-start sm:items-center gap-2.5 sm:gap-3 min-w-0">
                                <div class="w-8 h-8 rounded-xl bg-white/10 flex items-center justify-center font-black text-sm flex-shrink-0">2</div>
                                <div class="min-w-0">
                                    <h4 class="text-xs sm:text-sm md:text-base font-black text-white leading-snug">Core Doctor Maximization (1 Sweater / Doctor)</h4>
                                    <p class="text-[10px] sm:text-xs text-purple-100/80">3 Core Doctors per territory • 1 sweater per doctor</p>
                                </div>
                            </div>
                            <span id="c2-status-pill" class="px-2.5 py-1 rounded-full text-[10px] sm:text-xs font-bold bg-white/20 text-white backdrop-blur-sm self-start sm:self-auto flex-shrink-0">3 Doctors / Terr</span>
                        </div>

                        <div class="p-4 sm:p-6 space-y-4">
                            <!-- 3 Doctors Grid (2 Top side-by-side, 1 Bottom Center) -->
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
                                <!-- Doc 3 (Centered in Bottom Row) -->
                                <div class="bg-purple-50/50 border border-purple-200 rounded-2xl p-3.5 space-y-2.5 col-span-1 md:col-span-2 max-w-lg mx-auto w-full">
                                    <div class="flex items-center justify-between">
                                        <span class="text-xs font-bold text-purple-950 flex items-center gap-1.5"><span class="w-4 h-4 rounded-full bg-purple-600 text-white flex items-center justify-center text-[10px] font-black">3</span> Doctor 3</span>
                                        <span id="c2_d3_check_badge"><span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span></span>
                                    </div>
                                    <div class="space-y-2">
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

                    <!-- Bottom Nav / Action Bar -->
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
    <div id="lightbox-modal" class="fixed inset-0 z-50 bg-slate-950/90 backdrop-blur-md hidden flex items-center justify-center p-3 sm:p-6" onclick="closeLightbox()">
        <div class="bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl max-w-lg w-full overflow-hidden" onclick="event.stopPropagation()">
            <div class="p-4 border-b border-slate-800 flex items-center justify-between">
                <div>
                    <h3 id="lightbox-title" class="font-black text-sm text-white">Sweater Design Preview</h3>
                    <p id="lightbox-subtitle" class="text-xs text-orange-400">Design Details</p>
                </div>
                <button onclick="closeLightbox()" class="w-8 h-8 rounded-full bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center transition"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="p-4 space-y-4">
                <div class="aspect-[3/4] bg-slate-950 rounded-2xl overflow-hidden border border-slate-800 flex items-center justify-center relative">
                    <img id="lightbox-img" src="" alt="Sweater" class="w-full h-full object-cover">
                </div>
                <div id="lightbox-details" class="bg-slate-800/60 rounded-2xl p-3.5 text-xs space-y-1.5 border border-slate-700/50">
                    <!-- Populated dynamically -->
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

                <!-- DETAILED SIZE & MEASUREMENT SPECIFICATIONS -->
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
    <div id="admin-modal" class="fixed inset-0 z-50 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-3 sm:p-6" onclick="closeAdminModal()">
        <div class="bg-white border border-slate-200 rounded-3xl shadow-2xl max-w-5xl w-full max-h-[90vh] flex flex-col overflow-hidden text-slate-900" onclick="event.stopPropagation()">
            
            <!-- Modal Header -->
            <div class="px-5 py-4 border-b border-slate-200 flex items-center justify-between bg-slate-900 text-white">
                <div class="flex items-center gap-2.5">
                    <div class="w-8 h-8 rounded-xl bg-orange-500 flex items-center justify-center font-bold text-sm shadow-sm"><i class="fa-solid fa-shield-halved"></i></div>
                    <div>
                        <h3 class="font-bold text-xs sm:text-sm">Central Admin Campaign Dashboard</h3>
                        <p class="text-[10px] sm:text-xs text-slate-400">Campaign Oversight & Production Aggregator</p>
                    </div>
                </div>
                <button onclick="closeAdminModal()" class="w-8 h-8 rounded-full bg-slate-800 text-slate-400 hover:text-white flex items-center justify-center transition"><i class="fa-solid fa-xmark"></i></button>
            </div>

            <!-- Modal Body -->
            <div class="p-4 sm:p-6 overflow-y-auto custom-scrollbar space-y-6 flex-1 bg-slate-50">

                <!-- Admin Auth Card if Not Logged In -->
                <div id="admin-auth-card" class="bg-white border border-slate-200 rounded-2xl p-5 max-w-md mx-auto space-y-4 shadow-sm text-center">
                    <div class="w-12 h-12 rounded-2xl bg-orange-50 text-orange-600 flex items-center justify-center mx-auto text-xl border border-orange-100">
                        <i class="fa-solid fa-lock"></i>
                    </div>
                    <div>
                        <h4 class="text-base font-black text-slate-900">Admin Authentication</h4>
                        <p class="text-xs text-slate-500">Enter Admin PIN to view full live analytics and export master Excel</p>
                    </div>
                    <div class="space-y-3">
                        <input type="password" id="admin-password-input" onkeydown="if(event.key==='Enter') loginAdmin()" placeholder="Enter Admin PIN..." class="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-xs sm:text-sm text-slate-900 font-semibold focus:outline-none focus:border-orange-500 text-center tracking-widest">
                        <button onclick="loginAdmin()" class="w-full py-2.5 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-xl text-xs shadow transition active:scale-95">
                            Verify & Open Dashboard
                        </button>
                    </div>
                    <p class="text-[11px] text-slate-400">Default Admin PIN: <code class="font-mono font-bold text-slate-600">admin2026</code></p>
                </div>

                <!-- Admin Dashboard View (When Logged In) -->
                <div id="admin-dashboard-content" class="hidden space-y-6">

                    <!-- Actions & Master Export Bar -->
                    <div class="flex flex-wrap items-center justify-between gap-3 bg-white border border-slate-200 rounded-2xl p-4 shadow-sm">
                        <div class="flex flex-wrap items-center gap-2">
                            <button onclick="exportMasterExcelFromAdmin()" class="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white font-black rounded-xl text-xs flex items-center gap-2 shadow-sm transition active:scale-95">
                                <i class="fa-solid fa-file-excel text-sm"></i>
                                <span>Export Live Master Excel</span>
                            </button>
                            <button onclick="pullCloudData(true)" class="px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs border border-slate-300 flex items-center gap-1.5 transition active:scale-95">
                                <i class="fa-solid fa-cloud-arrow-down text-blue-600"></i>
                                <span>Pull Cloud Data</span>
                            </button>
                            <button onclick="toggleCloudSettings()" class="px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs border border-slate-300 flex items-center gap-1.5 transition active:scale-95">
                                <i class="fa-solid fa-gear text-slate-500"></i>
                                <span>Cloud Setup</span>
                            </button>
                        </div>
                        <div class="flex items-center gap-2">
                            <button onclick="deleteAllCampaignData()" class="px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold rounded-xl text-xs border border-rose-200 flex items-center gap-1.5 transition active:scale-95">
                                <i class="fa-solid fa-trash-can"></i>
                                <span>Delete All Data</span>
                            </button>
                        </div>
                    </div>

                    <!-- Cloud Settings Panel -->
                    <div id="admin-cloud-settings-box" class="hidden bg-slate-900 text-white rounded-2xl p-4 space-y-3 border border-slate-800">
                        <div class="flex items-center justify-between">
                            <div class="flex items-center gap-2">
                                <i class="fa-solid fa-brands fa-google-drive text-amber-400 text-lg"></i>
                                <span class="text-xs font-bold">Google Sheet & Apps Script Cloud Sync URL</span>
                            </div>
                            <button onclick="toggleCloudSettings()" class="text-slate-400 hover:text-white text-xs"><i class="fa-solid fa-xmark"></i></button>
                        </div>
                        <p class="text-[11px] text-slate-300">All submissions from Regional Managers across Bangladesh automatically sync in real-time to this Google Sheet Web App:</p>
                        <div class="flex gap-2">
                            <input type="text" id="custom-cloud-url-input" class="flex-1 bg-slate-800 border border-slate-700 rounded-xl px-3 py-1.5 text-xs text-slate-200 font-mono focus:outline-none focus:border-orange-500" placeholder="https://script.google.com/macros/s/.../exec">
                            <button onclick="saveCloudUrlSetting()" class="px-3 py-1.5 bg-orange-500 hover:bg-orange-600 text-white font-bold rounded-xl text-xs transition">Save</button>
                            <button onclick="testGoogleDriveConnection()" class="px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-slate-200 font-bold rounded-xl text-xs transition">Test</button>
                        </div>
                        <div id="cloud-test-result" class="text-[11px]"></div>
                    </div>

                    <!-- Top 4 KPI Cards -->
                    <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4">
                        <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                            <span class="text-[10px] font-bold uppercase text-slate-400 tracking-wider">Total Territories</span>
                            <div id="admin-kpi-total" class="text-xl sm:text-2xl font-black text-slate-900">1,856</div>
                            <p class="text-[10px] text-slate-500">Across 63 Regions in 19 Zones</p>
                        </div>
                        <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                            <span class="text-[10px] font-bold uppercase text-emerald-600 tracking-wider">Completed</span>
                            <div id="admin-kpi-complete" class="text-xl sm:text-2xl font-black text-emerald-600">0</div>
                            <p id="admin-kpi-complete-pct" class="text-[10px] text-emerald-700 font-bold">0.0% completion rate</p>
                        </div>
                        <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                            <span class="text-[10px] font-bold uppercase text-amber-600 tracking-wider">In Progress</span>
                            <div id="admin-kpi-progress" class="text-xl sm:text-2xl font-black text-amber-600">0</div>
                            <p class="text-[10px] text-slate-500">Partially filled</p>
                        </div>
                        <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm space-y-1">
                            <span class="text-[10px] font-bold uppercase text-orange-600 tracking-wider">Total Sweaters</span>
                            <div id="admin-kpi-sweaters" class="text-xl sm:text-2xl font-black text-orange-600">0</div>
                            <p class="text-[10px] text-slate-500">Aggregated for production</p>
                        </div>
                    </div>

                    <!-- Size Breakdown Matrix -->
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between">
                            <h4 class="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-wider flex items-center gap-2">
                                <i class="fa-solid fa-layer-group text-orange-500"></i>
                                <span>Sweater Production Breakdown Matrix</span>
                            </h4>
                            <span class="text-[10px] font-bold bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full">Live Aggregation</span>
                        </div>
                        <div class="overflow-x-auto">
                            <table class="w-full text-xs text-center border-collapse bg-white rounded-xl overflow-hidden border border-slate-200 shadow-sm">
                                <thead class="bg-slate-900 text-white text-[10px] font-bold uppercase">
                                    <tr>
                                        <th class="p-2.5 text-left">Item / Design</th>
                                        <th class="p-2.5">XS</th>
                                        <th class="p-2.5">S</th>
                                        <th class="p-2.5">M</th>
                                        <th class="p-2.5">L</th>
                                        <th class="p-2.5">XL</th>
                                        <th class="p-2.5">XXL</th>
                                        <th class="p-2.5 font-black text-orange-400">Total</th>
                                    </tr>
                                </thead>
                                <tbody id="admin-matrix-tbody" class="divide-y divide-slate-100 font-medium">
                                    <!-- Populated dynamically -->
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- Zone-wise Progress Summary -->
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
                        <h4 class="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-wider flex items-center gap-2">
                            <i class="fa-solid fa-chart-column text-teal-600"></i>
                            <span>Zone-wise Submission Progress</span>
                        </h4>
                        <div id="admin-zone-progress-list" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                            <!-- Populated dynamically -->
                        </div>
                    </div>

                    <!-- Regions List & Management -->
                    <div class="bg-white border border-slate-200 rounded-2xl p-4 sm:p-5 shadow-sm space-y-3">
                        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                            <h4 class="text-xs sm:text-sm font-black text-slate-900 uppercase tracking-wider flex items-center gap-2">
                                <i class="fa-solid fa-list-check text-purple-600"></i>
                                <span>All 63 Regions Detailed Status</span>
                            </h4>
                            <input type="text" id="admin-region-search" oninput="renderAdminRegionsTable(this.value)" placeholder="Search Region or Regional Head..." class="bg-slate-50 border border-slate-300 rounded-xl px-3 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-orange-500 w-full sm:w-64">
                        </div>
                        <div class="overflow-x-auto max-h-80 custom-scrollbar border border-slate-200 rounded-xl">
                            <table class="w-full text-xs text-left border-collapse bg-white">
                                <thead class="bg-slate-100 text-slate-700 text-[10px] font-bold uppercase sticky top-0 border-b border-slate-200">
                                    <tr>
                                        <th class="p-2.5">Region</th>
                                        <th class="p-2.5">Zone</th>
                                        <th class="p-2.5">Regional Head</th>
                                        <th class="p-2.5 text-center">Territories</th>
                                        <th class="p-2.5 text-center">Status</th>
                                        <th class="p-2.5 text-right">Action</th>
                                    </tr>
                                </thead>
                                <tbody id="admin-regions-tbody" class="divide-y divide-slate-100 font-medium text-slate-800">
                                    <!-- Populated dynamically -->
                                </tbody>
                            </table>
                        </div>
                    </div>

                </div>

            </div>
        </div>
    </div>

    <!-- TOAST NOTIFICATION -->
    <div id="toast" class="fixed bottom-5 right-5 z-50 transform translate-y-20 opacity-0 transition-all duration-300 pointer-events-none bg-slate-900 text-white border border-slate-800 px-4 py-3 rounded-2xl shadow-2xl flex items-center gap-2.5 text-xs font-semibold max-w-sm">
        <i class="fa-solid fa-circle-check text-emerald-400 text-sm"></i>
        <span id="toast-msg">Action completed</span>
    </div>

    <!-- JAVASCRIPT LOGIC -->
    <script>
        const ALL_TERRITORIES = ###ALL_TERRITORIES_JSON###;
        const REGION_MAP = ###REGION_MAP_JSON###;
        const ZONES = ###ZONES_JSON###;
        const DEFAULT_CLOUD_URL = "https://script.google.com/macros/s/AKfycbzEnDTtNiXEAyB5qHqrxLj1RbNytgOJAB_lKjw_VVVd1C8CiaeYU6iTROiJabkyX_-b/exec";

        const SWEATER_DETAILS = {
            "01": { code: "01", name: "Men's Sleeveless V-Neck Sweater", color: "Solid Ash / Grey Textured", gender: "Men's", sizes: ["S", "M", "L", "XL", "XXL"], img: "###B64_01###", supplier: "Richman / Lubnan" },
            "02": { code: "02", name: "Men's Sleeveless V-Neck Sweater", color: "Solid Navy Blue Textured", gender: "Men's", sizes: ["S", "M", "L", "XL", "XXL"], img: "###B64_02###", supplier: "Richman / Lubnan" },
            "03": { code: "03", name: "Men's Sleeveless V-Neck Sweater", color: "Off-White / Cream Check", gender: "Men's", sizes: ["S", "M", "L", "XL", "XXL"], img: "###B64_03###", supplier: "Richman / Lubnan" },
            "04": { code: "04", name: "Women's Short Cardigan", color: "White & Navy Grid Check", gender: "Women's", sizes: ["XS", "S", "M", "L", "XL"], img: "###B64_04###", supplier: "Richman / Lubnan" },
            "05": { code: "05", name: "Women's Semi Long Cardigan", color: "Solid Black with Border Trim", gender: "Women's", sizes: ["S", "M", "L", "XL", "XXL"], img: "###B64_05###", supplier: "Richman / Lubnan" }
        };

        let store = JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{}');
        let regionLocks = JSON.parse(localStorage.getItem('EXIUM_REGION_LOCKS') || '{}');
        let isGlobalAccessOpen = JSON.parse(localStorage.getItem('EXIUM_GLOBAL_ACCESS') || 'true');
        let cloudApiUrl = localStorage.getItem('EXIUM_CLOUD_URL') || DEFAULT_CLOUD_URL;

        let currentRegionCode = null;
        let activeTerritoryIndex = 0;
        let isAdminLoggedIn = false;
        let autoSyncTimeout = null;

        window.addEventListener('DOMContentLoaded', () => {
            populateZoneDropdown();
            checkGlobalLockBanner();

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

            regSel.innerHTML = '<option value="">-- Select Region --</option>';
            rhCard.classList.add('hidden');
            passCard.classList.add('hidden');
            unlockBtn.classList.add('hidden');

            if (!zone) {
                regSel.disabled = true;
                return;
            }

            const matchingRegions = Object.values(REGION_MAP).filter(r => r.zone === zone);
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

            if (!regCode || !REGION_MAP[regCode]) {
                rhCard.classList.add('hidden');
                passCard.classList.add('hidden');
                unlockBtn.classList.add('hidden');
                return;
            }

            const r = REGION_MAP[regCode];
            document.getElementById('rh-name-display').innerHTML = `<i class="fa-solid fa-user-tie text-orange-500"></i> <span>${r.regional_head}</span>`;
            document.getElementById('rh-territory-count').innerHTML = `Total Territories: <strong>${r.territories.length}</strong>`;
            rhCard.classList.remove('hidden');
            passCard.classList.remove('hidden');
            unlockBtn.classList.remove('hidden');

            if (passInput) {
                passInput.value = '';
                passInput.focus();
            }
        }

        function handlePasswordKey(e) {
            if (e.key === 'Enter') unlockRegion();
        }

        function unlockRegion(bypassCode = null, isRestoringSession = false) {
            const code = bypassCode || document.getElementById('select-region').value;
            const passInput = document.getElementById('region-password');
            const pass = passInput ? passInput.value.trim() : '';

            if (!code || !REGION_MAP[code]) {
                showToast("⚠️ Please select a valid Region.");
                return;
            }

            if (!bypassCode) {
                if (pass !== '1234' && pass !== 'admin2026') {
                    showToast("❌ Incorrect Region Password (Default: 1234)");
                    return;
                }
            }

            const r = REGION_MAP[code];
            currentRegionCode = code;
            activeTerritoryIndex = 0;

            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify({
                region_code: code,
                territory_idx: activeTerritoryIndex
            }));

            document.getElementById('view-login').classList.add('hidden');
            document.getElementById('view-workspace').classList.remove('hidden');

            document.getElementById('banner-zone').textContent = r.zone;
            document.getElementById('banner-region').textContent = `SAP: ${r.sap_region_code}`;
            document.getElementById('banner-rh').textContent = `Region: ${r.region_name} (${r.regional_head})`;
            document.getElementById('banner-total-count').textContent = r.territories.length;

            renderTerritoryTabs();
            selectTerritoryTab(0, false);
            updateRegionCompletionBanner();

            if (!isRestoringSession) {
                showToast(`👋 Welcome ${r.regional_head}! Workspace loaded.`);
            }
        }

        function exitRegionWorkspace() {
            currentRegionCode = null;
            activeTerritoryIndex = 0;
            localStorage.removeItem('EXIUM_ACTIVE_SESSION');

            document.getElementById('view-workspace').classList.add('hidden');
            document.getElementById('view-login').classList.remove('hidden');
            const passInput = document.getElementById('region-password');
            if (passInput) passInput.value = '';
        }

        function renderTerritoryTabs() {
            const r = REGION_MAP[currentRegionCode];
            const deskList = document.getElementById('desktop-territory-list');
            const mobSelect = document.getElementById('mobile-territory-select');

            deskList.innerHTML = '';
            mobSelect.innerHTML = '';

            r.territories.forEach((t, idx) => {
                const terrCode = String(t.sap_territory_code);
                const d = store[terrCode] || {};
                const st = getTerritoryStatus(d);

                let badgeHtml = '<span class="text-slate-400 text-[10px]"><i class="fa-regular fa-circle"></i></span>';
                if (st === 'Complete') {
                    badgeHtml = '<span class="text-emerald-600 text-[11px] font-bold"><i class="fa-solid fa-circle-check"></i></span>';
                } else if (st === 'In Progress') {
                    badgeHtml = '<span class="text-amber-500 text-[11px] font-bold"><i class="fa-solid fa-circle-half-stroke"></i></span>';
                }

                const btn = document.createElement('button');
                btn.id = `desk-terr-tab-${idx}`;
                btn.onclick = () => selectTerritoryTab(idx);
                btn.className = `w-full text-left px-3 py-2.5 rounded-xl border text-xs font-semibold flex items-center justify-between transition ${
                    idx === activeTerritoryIndex 
                        ? 'nav-tab-active bg-orange-500 text-white border-orange-500 shadow-sm' 
                        : 'bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200'
                }`;
                btn.innerHTML = `
                    <div class="truncate pr-1">
                        <span class="font-bold block truncate">${t.territory_name}</span>
                        <span class="text-[10px] opacity-75 font-mono">SAP: ${t.sap_territory_code}</span>
                    </div>
                    <div>${badgeHtml}</div>
                `;
                deskList.appendChild(btn);

                const opt = document.createElement('option');
                opt.value = idx;
                opt.textContent = `${idx + 1}. ${t.territory_name} (${st})`;
                mobSelect.appendChild(opt);
            });

            mobSelect.value = activeTerritoryIndex;
        }

        function selectTerritoryTab(idx, saveCurrent = true) {
            if (saveCurrent) onDataChanged();

            const r = REGION_MAP[currentRegionCode];
            if (!r || !r.territories[idx]) return;

            activeTerritoryIndex = idx;

            const savedSession = JSON.parse(localStorage.getItem('EXIUM_ACTIVE_SESSION') || '{}');
            savedSession.territory_idx = idx;
            localStorage.setItem('EXIUM_ACTIVE_SESSION', JSON.stringify(savedSession));

            const mobSelect = document.getElementById('mobile-territory-select');
            if (mobSelect) mobSelect.value = idx;

            r.territories.forEach((t, i) => {
                const tab = document.getElementById(`desk-terr-tab-${i}`);
                if (tab) {
                    if (i === idx) {
                        tab.className = 'w-full text-left px-3 py-2.5 rounded-xl border text-xs font-bold flex items-center justify-between transition nav-tab-active bg-orange-500 text-white border-orange-500 shadow-sm';
                    } else {
                        tab.className = 'w-full text-left px-3 py-2.5 rounded-xl border text-xs font-semibold flex items-center justify-between transition bg-slate-50 hover:bg-slate-100 text-slate-700 border-slate-200';
                    }
                }
            });

            const t = r.territories[idx];
            renderTerritoryData(t);
        }

        function navigateTerritory(dir) {
            onDataChanged();
            const r = REGION_MAP[currentRegionCode];
            if (!r) return;
            let target = activeTerritoryIndex + dir;
            if (target >= 0 && target < r.territories.length) {
                selectTerritoryTab(target);
            } else if (target >= r.territories.length) {
                showToast("🏁 Reached last territory in this Region!");
            }
        }

        function renderTerritoryData(t) {
            const terrCode = String(t.sap_territory_code);
            const d = store[terrCode] || {};

            document.getElementById('active-terr-sap').textContent = `SAP Territory Code: ${t.sap_territory_code}`;
            document.getElementById('active-terr-name').textContent = `${t.territory_name} (${t.zone})`;

            const isLocked = isRegionLocked();
            const lockedNotice = document.getElementById('territory-locked-notice');
            if (isLocked) {
                lockedNotice.classList.remove('hidden');
            } else {
                lockedNotice.classList.add('hidden');
            }

            // Campaign 1: Doctor Name & RPL
            const c1DocInput = document.getElementById('c1_doc_name');
            c1DocInput.value = d.c1_doc_name || '';
            c1DocInput.disabled = isLocked;

            const c1DocRpl = document.getElementById('c1_doc_rpl');
            c1DocRpl.value = d.c1_doc_rpl || '';
            c1DocRpl.disabled = isLocked;
            updateRplBadgeState(c1DocRpl, 'c1_doc_rpl_badge');

            // Campaign 1: Sweaters 1 to 4
            ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                const sw = d[`c1_${m}_sweater`] || '';
                const sz = d[`c1_${m}_size`] || '';
                const swSel = document.getElementById(`c1_${m}_sweater`);
                const szSel = document.getElementById(`c1_${m}_size`);
                
                swSel.value = sw;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c1_${m}_sweater`, `c1_${m}_size`, sz);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c1_${m}`, sw);
            });

            // Campaign 1: Sweater 4 Visibility
            if (d.c1_m4_sweater || d.c1_m4_size) {
                showC1Sweater4(false);
            } else {
                hideC1Sweater4ViewOnly();
            }

            // Campaign 2: Doctors 1, 2, 3
            ['d1', 'd2', 'd3'].forEach(dSlot => {
                const nameInp = document.getElementById(`c2_${dSlot}_name`);
                const rplInp = document.getElementById(`c2_${dSlot}_rpl`);
                const swSel = document.getElementById(`c2_${dSlot}_sweater`);
                const szSel = document.getElementById(`c2_${dSlot}_size`);

                const nameVal = d[`c2_${dSlot}_name`] || '';
                const rplVal = d[`c2_${dSlot}_rpl`] || '';
                const swVal = d[`c2_${dSlot}_sweater`] || '';
                const szVal = d[`c2_${dSlot}_size`] || '';

                nameInp.value = nameVal;
                nameInp.disabled = isLocked;

                rplInp.value = rplVal;
                rplInp.disabled = isLocked;
                updateRplBadgeState(rplInp, `c2_${dSlot}_rpl_badge`);

                swSel.value = swVal;
                swSel.disabled = isLocked;
                updateSizeOptionsForSelect(`c2_${dSlot}_sweater`, `c2_${dSlot}_size`, szVal);
                szSel.disabled = isLocked;
                updateSlotImagePreview(`c2_${dSlot}`, swVal);
            });

            updateTerritoryCardBadges(d);
            updateActiveTerritoryStatusBadge(d);
        }

        function showC1Sweater4(triggerChange = false) {
            document.getElementById('c1_m4_add_container').classList.add('hidden');
            document.getElementById('c1_m4_card').classList.remove('hidden');
            if (triggerChange) onDataChanged();
        }

        function hideC1Sweater4ViewOnly() {
            document.getElementById('c1_m4_card').classList.add('hidden');
            document.getElementById('c1_m4_add_container').classList.remove('hidden');
        }

        function hideAndClearC1Sweater4() {
            document.getElementById('c1_m4_sweater').value = '';
            document.getElementById('c1_m4_size').innerHTML = '<option value="">-- Size --</option>';
            document.getElementById('c1_m4_size').value = '';
            updateSlotImagePreview('c1_m4', '');
            hideC1Sweater4ViewOnly();
            onDataChanged();
            showToast("ℹ️ Sweater 4 option removed.");
        }

        function updateRplBadgeState(input, badgeId) {
            const b = document.getElementById(badgeId);
            if (!b) return;
            const val = input.value.trim();
            if (val.length === 6 && /^[0-9]{6}$/.test(val)) {
                b.innerHTML = '<i class="fa-solid fa-check text-emerald-600"></i> Valid RPL';
                b.className = 'text-[9px] font-bold text-emerald-600';
            } else if (val.length > 0) {
                b.innerHTML = `${val.length}/6 digits`;
                b.className = 'text-[9px] font-bold text-amber-500';
            } else {
                b.innerHTML = '6 digits';
                b.className = 'text-[9px] font-bold text-slate-400';
            }
        }

        function onRplInput(input, badgeId) {
            input.value = input.value.replace(/[^0-9]/g, '').slice(0, 6);
            updateRplBadgeState(input, badgeId);
            onDataChanged();
        }

        function updateSizeOptionsForSelect(sweaterSelectId, sizeSelectId, selectedSize = '') {
            const swVal = document.getElementById(sweaterSelectId)?.value || '';
            const szSel = document.getElementById(sizeSelectId);
            if (!szSel) return;

            szSel.innerHTML = '<option value="">-- Size --</option>';
            if (!swVal) return;

            const code = swVal.split(' ')[0];
            const details = SWEATER_DETAILS[code];
            if (details && details.sizes) {
                details.sizes.forEach(sz => {
                    const opt = document.createElement('option');
                    opt.value = sz;
                    opt.textContent = sz;
                    if (sz === selectedSize) opt.selected = true;
                    szSel.appendChild(opt);
                });
            }
        }

        function onSweaterSelectChange(prefix, sweaterVal) {
            updateSizeOptionsForSelect(`${prefix}_sweater`, `${prefix}_size`, '');
            updateSlotImagePreview(prefix, sweaterVal);
            onDataChanged();
        }

        function updateSlotImagePreview(prefix, sweaterVal) {
            const prev = document.getElementById(`${prefix}_img_preview`);
            if (!prev) return;

            if (!sweaterVal) {
                prev.innerHTML = '<i class="fa-solid fa-shirt text-lg text-slate-300"></i>';
                return;
            }

            const code = sweaterVal.split(' ')[0];
            const details = SWEATER_DETAILS[code];
            if (details && details.img) {
                prev.innerHTML = `<img src="${details.img}" alt="${details.name}" class="w-full h-full object-cover">`;
            } else {
                prev.innerHTML = '<i class="fa-solid fa-shirt text-lg text-orange-400"></i>';
            }
        }

        function onDataChanged() {
            if (isRegionLocked() || !currentRegionCode) return;
            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[activeTerritoryIndex];
            const terrCode = String(t.sap_territory_code);

            const isC1M4Visible = !document.getElementById('c1_m4_card').classList.contains('hidden');

            const terrData = {
                c1_doc_name: document.getElementById('c1_doc_name').value.trim(),
                c1_doc_rpl: document.getElementById('c1_doc_rpl').value.trim(),
                c1_m1_sweater: document.getElementById('c1_m1_sweater').value,
                c1_m1_size: document.getElementById('c1_m1_size').value,
                c1_m2_sweater: document.getElementById('c1_m2_sweater').value,
                c1_m2_size: document.getElementById('c1_m2_size').value,
                c1_m3_sweater: document.getElementById('c1_m3_sweater').value,
                c1_m3_size: document.getElementById('c1_m3_size').value,
                c1_m4_sweater: isC1M4Visible ? document.getElementById('c1_m4_sweater').value : '',
                c1_m4_size: isC1M4Visible ? document.getElementById('c1_m4_size').value : '',

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

                c2_d4_name: '',
                c2_d4_rpl: '',
                c2_d4_sweater: '',
                c2_d4_size: ''
            };

            store[terrCode] = terrData;
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));

            updateTerritoryCardBadges(terrData);
            updateActiveTerritoryStatusBadge(terrData);
            updateTabBadge(activeTerritoryIndex, terrData);
            updateRegionCompletionBanner();

            if (autoSyncTimeout) clearTimeout(autoSyncTimeout);
            autoSyncTimeout = setTimeout(() => {
                syncTerritoryToCloud(terrCode, terrData);
            }, 1200);
        }

        function updateSweaterSlotBadge(slotId, sweaterVal, sizeVal) {
            const b = document.getElementById(`${slotId}_check_badge`);
            if (!b) return;
            if (sweaterVal && sizeVal) {
                b.innerHTML = '<span class="text-emerald-600 text-[10px] font-bold flex items-center gap-1"><i class="fa-solid fa-check"></i> Selected</span>';
            } else if (sweaterVal || sizeVal) {
                b.innerHTML = '<span class="text-amber-500 text-[10px] font-bold">Incomplete</span>';
            } else {
                b.innerHTML = '<span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span>';
            }
        }

        function updateDoctorCardBadge(docSlot, nameVal, rplVal, sweaterVal, sizeVal) {
            const b = document.getElementById(`c2_${docSlot}_check_badge`);
            if (!b) return;
            const isComp = nameVal && rplVal && String(rplVal).length === 6 && sweaterVal && sizeVal;
            if (isComp) {
                b.innerHTML = '<span class="text-emerald-600 text-[10px] font-bold flex items-center gap-1"><i class="fa-solid fa-check"></i> Complete</span>';
            } else if (nameVal || rplVal || sweaterVal || sizeVal) {
                b.innerHTML = '<span class="text-amber-500 text-[10px] font-bold">Incomplete</span>';
            } else {
                b.innerHTML = '<span class="text-slate-400 text-[10px] font-medium"><i class="fa-regular fa-circle"></i> Pending</span>';
            }
        }

        function updateTerritoryCardBadges(d) {
            updateSweaterSlotBadge('c1_m1', d.c1_m1_sweater, d.c1_m1_size);
            updateSweaterSlotBadge('c1_m2', d.c1_m2_sweater, d.c1_m2_size);
            updateSweaterSlotBadge('c1_m3', d.c1_m3_sweater, d.c1_m3_size);
            updateSweaterSlotBadge('c1_m4', d.c1_m4_sweater, d.c1_m4_size);

            updateDoctorCardBadge('d1', d.c2_d1_name, d.c2_d1_rpl, d.c2_d1_sweater, d.c2_d1_size);
            updateDoctorCardBadge('d2', d.c2_d2_name, d.c2_d2_rpl, d.c2_d2_sweater, d.c2_d2_size);
            updateDoctorCardBadge('d3', d.c2_d3_name, d.c2_d3_rpl, d.c2_d3_sweater, d.c2_d3_size);
        }

        function getTerritoryStatus(d) {
            if (!d) return 'Not Started';

            // Campaign 1: Valid if Doctor Name, 6-digit RPL, Sw1, Sw2, Sw3 are all chosen.
            // If Sw4 is chosen, then Sw4 size must also be chosen.
            const c1Ok = d.c1_doc_name && d.c1_doc_rpl && String(d.c1_doc_rpl).length === 6 &&
                         d.c1_m1_sweater && d.c1_m1_size &&
                         d.c1_m2_sweater && d.c1_m2_size &&
                         d.c1_m3_sweater && d.c1_m3_size &&
                         (!d.c1_m4_sweater || (d.c1_m4_sweater && d.c1_m4_size));

            // Campaign 2: Valid if Doctor 1, 2, 3 are all completely filled
            const c2Ok = d.c2_d1_name && d.c2_d1_rpl && String(d.c2_d1_rpl).length === 6 && d.c2_d1_sweater && d.c2_d1_size &&
                         d.c2_d2_name && d.c2_d2_rpl && String(d.c2_d2_rpl).length === 6 && d.c2_d2_sweater && d.c2_d2_size &&
                         d.c2_d3_name && d.c2_d3_rpl && String(d.c2_d3_rpl).length === 6 && d.c2_d3_sweater && d.c2_d3_size;

            if (c1Ok && c2Ok) return 'Complete';

            if (d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater || d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater || d.c2_d2_name || d.c2_d3_name) {
                return 'In Progress';
            }
            return 'Not Started';
        }

        function updateActiveTerritoryStatusBadge(d) {
            const st = getTerritoryStatus(d);
            const badge = document.getElementById('active-terr-status');
            if (!badge) return;

            if (st === 'Complete') {
                badge.className = 'px-3 py-1 rounded-full text-xs font-black bg-emerald-100 text-emerald-800 border border-emerald-300 flex items-center gap-1.5';
                badge.innerHTML = '<i class="fa-solid fa-circle-check text-emerald-600"></i> <span>Complete</span>';
            } else if (st === 'In Progress') {
                badge.className = 'px-3 py-1 rounded-full text-xs font-black bg-amber-100 text-amber-800 border border-amber-300 flex items-center gap-1.5';
                badge.innerHTML = '<i class="fa-solid fa-circle-half-stroke text-amber-600"></i> <span>In Progress</span>';
            } else {
                badge.className = 'px-3 py-1 rounded-full text-xs font-black bg-slate-100 text-slate-600 border border-slate-200 flex items-center gap-1.5';
                badge.innerHTML = '<i class="fa-regular fa-circle text-[9px]"></i> <span>Not Started</span>';
            }
        }

        function updateTabBadge(idx, d) {
            const st = getTerritoryStatus(d);
            const tab = document.getElementById(`desk-terr-tab-${idx}`);
            if (tab) {
                const iconContainer = tab.querySelector('div:last-child');
                if (iconContainer) {
                    if (st === 'Complete') {
                        iconContainer.innerHTML = '<span class="text-emerald-600 text-[11px] font-bold"><i class="fa-solid fa-circle-check"></i></span>';
                    } else if (st === 'In Progress') {
                        iconContainer.innerHTML = '<span class="text-amber-500 text-[11px] font-bold"><i class="fa-solid fa-circle-half-stroke"></i></span>';
                    } else {
                        iconContainer.innerHTML = '<span class="text-slate-400 text-[10px]"><i class="fa-regular fa-circle"></i></span>';
                    }
                }
            }
            const mobSelect = document.getElementById('mobile-territory-select');
            if (mobSelect && mobSelect.options[idx]) {
                const r = REGION_MAP[currentRegionCode];
                const tName = r.territories[idx].territory_name;
                mobSelect.options[idx].textContent = `${idx + 1}. ${tName} (${st})`;
            }
        }

        function updateRegionCompletionBanner() {
            const r = REGION_MAP[currentRegionCode];
            if (!r) return;
            let compCount = 0;
            r.territories.forEach(t => {
                const st = getTerritoryStatus(store[String(t.sap_territory_code)]);
                if (st === 'Complete') compCount++;
            });
            document.getElementById('banner-complete-count').textContent = compCount;
            document.getElementById('region-progress-badge').textContent = `${compCount}/${r.territories.length} Done`;
        }

        function saveCurrentTerritoryClick() {
            onDataChanged();
            const r = REGION_MAP[currentRegionCode];
            const t = r.territories[activeTerritoryIndex];
            const terrCode = String(t.sap_territory_code);
            const d = store[terrCode] || {};
            const st = getTerritoryStatus(d);

            syncTerritoryToCloud(terrCode, d);

            if (st === 'Complete') {
                showToast(`✅ Saved! ${t.territory_name} is 100% Complete.`);
            } else {
                showToast(`💾 Saved ${t.territory_name} (${st}).`);
            }
        }

        function isRegionLocked() {
            if (!isGlobalAccessOpen) return true;
            if (currentRegionCode && regionLocks[currentRegionCode]) return true;
            return false;
        }

        function checkGlobalLockBanner() {
            const b = document.getElementById('login-global-locked-alert');
            if (b) {
                if (!isGlobalAccessOpen) b.classList.remove('hidden');
                else b.classList.add('hidden');
            }
        }

        function togglePasswordVisibility(inputId, btn) {
            const inp = document.getElementById(inputId);
            const icon = btn.querySelector('i');
            if (inp.type === 'password') {
                inp.type = 'text';
                icon.className = 'fa-regular fa-eye-slash';
            } else {
                inp.type = 'password';
                icon.className = 'fa-regular fa-eye';
            }
        }

        function toggleCloudSettings() {
            const box = document.getElementById('admin-cloud-settings-box');
            if (!box) return;
            box.classList.toggle('hidden');
            const input = document.getElementById('custom-cloud-url-input');
            if (input) input.value = cloudApiUrl;
        }

        function saveCloudUrlSetting() {
            const url = document.getElementById('custom-cloud-url-input').value.trim();
            cloudApiUrl = url || DEFAULT_CLOUD_URL;
            localStorage.setItem('EXIUM_CLOUD_URL', cloudApiUrl);
            showToast("💾 Cloud URL configuration saved!");
        }

        async function testGoogleDriveConnection() {
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (!url) {
                showCloudTestResult("No Google Apps Script URL set", "text-rose-400");
                return;
            }
            showCloudTestResult("Pinging Google Cloud...", "text-amber-400");
            try {
                const res = await fetch(url + (url.includes('?') ? '&' : '?') + 'action=ping');
                const json = await res.json();
                if (json && json.status === 'success') {
                    showCloudTestResult("✅ Connected! Google Drive & Sheet active.", "text-emerald-400 font-bold");
                } else {
                    showCloudTestResult("⚠️ Unexpected response from server.", "text-amber-400");
                }
            } catch (err) {
                showCloudTestResult("❌ Connection Failed. Check URL deployment.", "text-rose-400");
            }
        }

        function showCloudTestResult(msg, className) {
            const el = document.getElementById('cloud-test-result');
            if (el) {
                el.className = className + " mt-2";
                el.textContent = msg;
            }
        }

        async function syncTerritoryToCloud(terrCode, terrData) {
            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (!url) return;

            try {
                const payload = {
                    action: "save_territory",
                    sap_territory_code: terrCode,
                    data: terrData
                };
                await fetch(url, {
                    method: "POST",
                    mode: "no-cors",
                    headers: { "Content-Type": "application/json" },
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
                    showToast("⚠️ Could not pull cloud data. Please verify connection.");
                }
                return { success: false, count: 0 };
            }
        }

        function loginAdmin() {
            const p = document.getElementById('admin-password-input').value.trim();
            if (p === 'admin2026') {
                isAdminLoggedIn = true;
                document.getElementById('admin-auth-card').classList.add('hidden');
                document.getElementById('admin-dashboard-content').classList.remove('hidden');
                
                renderAdminKpisAndSummaries();
                renderAdminZoneProgress();
                renderAdminProductionMatrix();
                renderAdminRegionsTable('');
                showToast("🔓 Admin authenticated successfully!");
            } else {
                showToast("❌ Incorrect Admin PIN (Default: admin2026)");
            }
        }

        function openAdminModal() {
            document.getElementById('admin-modal').classList.remove('hidden');
            if (!isAdminLoggedIn) {
                document.getElementById('admin-auth-card').classList.remove('hidden');
                document.getElementById('admin-dashboard-content').classList.add('hidden');
                const inp = document.getElementById('admin-password-input');
                if (inp) { inp.value = ''; inp.focus(); }
            } else {
                renderAdminKpisAndSummaries();
                renderAdminZoneProgress();
                renderAdminProductionMatrix();
                renderAdminRegionsTable(document.getElementById('admin-region-search')?.value || '');
            }
        }

        function closeAdminModal() {
            document.getElementById('admin-modal').classList.add('hidden');
        }

        function renderAdminKpisAndSummaries() {
            let total = ALL_TERRITORIES.length;
            let comp = 0;
            let inProg = 0;

            ALL_TERRITORIES.forEach(t => {
                const st = getTerritoryStatus(store[String(t['SAP Territory Code'])]);
                if (st === 'Complete') comp++;
                else if (st === 'In Progress') inProg++;
            });

            const totals = calculateProductionTotals();
            const pct = total > 0 ? ((comp / total) * 100).toFixed(1) : 0;

            document.getElementById('admin-kpi-total').textContent = total.toLocaleString();
            document.getElementById('admin-kpi-complete').textContent = comp.toLocaleString();
            document.getElementById('admin-kpi-complete-pct').textContent = `${pct}% completion rate`;
            document.getElementById('admin-kpi-progress').textContent = inProg.toLocaleString();
            document.getElementById('admin-kpi-sweaters').textContent = totals.grandTotal.toLocaleString();
        }

        function calculateProductionTotals() {
            const counts = {
                "01": { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "02": { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "03": { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "04": { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 },
                "05": { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 }
            };
            let grandTotal = 0;

            ALL_TERRITORIES.forEach(t => {
                const terrCode = String(t['SAP Territory Code']);
                const d = store[terrCode] || {};

                // Campaign 1: Sweaters 1 to 4
                ['m1', 'm2', 'm3', 'm4'].forEach(m => {
                    const sw = d[`c1_${m}_sweater`];
                    const sz = d[`c1_${m}_size`];
                    if (sw && sz) {
                        const code = sw.split(' ')[0];
                        if (counts[code] && counts[code][sz] !== undefined) {
                            counts[code][sz]++;
                            counts[code].total++;
                            grandTotal++;
                        }}
                });

                // Campaign 2: Doctors 1 to 3
                ['d1', 'd2', 'd3'].forEach(dSlot => {
                    const sw = d[`c2_${dSlot}_sweater`];
                    const sz = d[`c2_${dSlot}_size`];
                    if (sw && sz) {
                        const code = sw.split(' ')[0];
                        if (counts[code] && counts[code][sz] !== undefined) {
                            counts[code][sz]++;
                            counts[code].total++;
                            grandTotal++;
                        }}
                });
            });

            return { counts, grandTotal };
        }

        function renderAdminProductionMatrix() {
            const { counts, grandTotal } = calculateProductionTotals();
            const tbody = document.getElementById('admin-matrix-tbody');
            if (!tbody) return;

            tbody.innerHTML = '';
            const sizeList = ["XS", "S", "M", "L", "XL", "XXL"];
            const sizeTotals = { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0 };

            for (let code in SWEATER_DETAILS) {
                const det = SWEATER_DETAILS[code];
                const c = counts[code] || { XS: 0, S: 0, M: 0, L: 0, XL: 0, XXL: 0, total: 0 };

                sizeList.forEach(sz => { sizeTotals[sz] += c[sz] || 0; });

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="p-2.5 text-left font-bold text-slate-900 flex items-center gap-2">
                        <span class="w-5 h-5 rounded bg-slate-100 text-slate-700 text-[10px] font-black flex items-center justify-center border border-slate-300">${code}</span>
                        <div class="truncate">
                            <span class="font-bold">${det.name}</span>
                            <span class="text-[10px] text-slate-400 block">${det.color}</span>
                        </div>
                    </td>
                    <td class="p-2.5 ${c.XS > 0 ? 'font-bold text-slate-900' : 'text-slate-300'}">${c.XS}</td>
                    <td class="p-2.5 ${c.S > 0 ? 'font-bold text-slate-900' : 'text-slate-300'}">${c.S}</td>
                    <td class="p-2.5 ${c.M > 0 ? 'font-bold text-slate-900' : 'text-slate-300'}">${c.M}</td>
                    <td class="p-2.5 ${c.L > 0 ? 'font-bold text-slate-900' : 'text-slate-300'}">${c.L}</td>
                    <td class="p-2.5 ${c.XL > 0 ? 'font-bold text-slate-900' : 'text-slate-300'}">${c.XL}</td>
                    <td class="p-2.5 ${c.XXL > 0 ? 'font-bold text-slate-900' : 'text-slate-300'}">${c.XXL}</td>
                    <td class="p-2.5 font-black text-orange-600 bg-orange-50/50">${c.total}</td>
                `;
                tbody.appendChild(tr);
            }

            const footerTr = document.createElement('tr');
            footerTr.className = 'bg-slate-900 text-white font-bold text-[11px]';
            footerTr.innerHTML = `
                <td class="p-2.5 text-left uppercase">Total Units</td>
                <td class="p-2.5">${sizeTotals.XS}</td>
                <td class="p-2.5">${sizeTotals.S}</td>
                <td class="p-2.5">${sizeTotals.M}</td>
                <td class="p-2.5">${sizeTotals.L}</td>
                <td class="p-2.5">${sizeTotals.XL}</td>
                <td class="p-2.5">${sizeTotals.XXL}</td>
                <td class="p-2.5 font-black text-orange-400 bg-slate-800">${grandTotal}</td>
            `;
            tbody.appendChild(footerTr);
        }

        function renderAdminZoneProgress() {
            const list = document.getElementById('admin-zone-progress-list');
            if (!list) return;
            list.innerHTML = '';

            ZONES.forEach(z => {
                const zoneTerritories = ALL_TERRITORIES.filter(t => t.Zone === z);
                let comp = 0;
                zoneTerritories.forEach(t => {
                    const st = getTerritoryStatus(store[String(t['SAP Territory Code'])]);
                    if (st === 'Complete') comp++;
                });

                const total = zoneTerritories.length;
                const pct = total > 0 ? Math.round((comp / total) * 100) : 0;

                const card = document.createElement('div');
                card.className = 'bg-slate-50 border border-slate-200 rounded-xl p-3 space-y-2';
                card.innerHTML = `
                    <div class="flex items-center justify-between text-xs font-bold text-slate-900">
                        <span class="truncate pr-1">${z}</span>
                        <span class="text-[10px] font-black px-2 py-0.5 rounded-full ${pct === 100 ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-200 text-slate-700'}">${comp}/${total} (${pct}%)</span>
                    </div>
                    <div class="w-full bg-slate-200 rounded-full h-2 overflow-hidden">
                        <div class="bg-gradient-to-r from-orange-500 to-amber-500 h-2 rounded-full transition-all duration-500" style="width: ${pct}%"></div>
                    </div>
                `;
                list.appendChild(card);
            });
        }

        function renderAdminRegionsTable(searchTerm = '') {
            const tbody = document.getElementById('admin-regions-tbody');
            if (!tbody) return;
            tbody.innerHTML = '';

            const term = (searchTerm || '').toLowerCase().trim();
            const regions = Object.values(REGION_MAP);

            regions.forEach(r => {
                if (term) {
                    const matchName = r.region_name.toLowerCase().includes(term);
                    const matchHead = r.regional_head.toLowerCase().includes(term);
                    const matchZone = r.zone.toLowerCase().includes(term);
                    const matchCode = r.sap_region_code.toLowerCase().includes(term);
                    if (!matchName && !matchHead && !matchZone && !matchCode) return;
                }

                let comp = 0;
                r.territories.forEach(t => {
                    const st = getTerritoryStatus(store[String(t.sap_territory_code)]);
                    if (st === 'Complete') comp++;
                });
                const total = r.territories.length;
                const isAllDone = total > 0 && comp === total;

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="p-2.5 font-bold text-slate-900">
                        <div>${r.region_name}</div>
                        <span class="text-[10px] text-slate-400 font-mono">SAP: ${r.sap_region_code}</span>
                    </td>
                    <td class="p-2.5 text-slate-600">${r.zone}</td>
                    <td class="p-2.5 text-slate-800 font-semibold">${r.regional_head}</td>
                    <td class="p-2.5 text-center font-bold">${comp} / ${total}</td>
                    <td class="p-2.5 text-center">
                        <span class="px-2 py-0.5 rounded-full text-[10px] font-bold ${
                            isAllDone 
                                ? 'bg-emerald-100 text-emerald-800' 
                                : (comp > 0 ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-500')
                        }">
                            ${isAllDone ? '100% Done' : (comp > 0 ? 'In Progress' : 'Not Started')}
                        </span>
                    </td>
                    <td class="p-2.5 text-right space-x-1">
                        <button onclick="adminQuickUnlock('${r.sap_region_code}')" class="px-2.5 py-1 bg-orange-50 hover:bg-orange-100 text-orange-700 font-bold rounded-lg text-[10px] transition" title="Open Workspace">
                            Open
                        </button>
                        <button onclick="deleteSingleRegionData('${r.sap_region_code}')" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-600 font-bold rounded-lg text-[10px] transition" title="Clear Region Data">
                            Clear
                        </button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function adminQuickUnlock(regCode) {
            closeAdminModal();
            unlockRegion(regCode, false);
        }

        async function deleteSingleRegionData(regCode) {
            const r = REGION_MAP[regCode];
            if (!r) return;
            if (!confirm(`⚠️ Are you sure you want to completely clear all data for ${r.region_name} (${regCode})?\\n\\nThis will reset both the portal and the Google Sheet.`)) return;

            r.territories.forEach(t => {
                delete store[String(t.sap_territory_code)];
            });
            localStorage.setItem('EXIUM_SWEATER_STORE', JSON.stringify(store));

            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (url) {
                try {
                    await fetch(url, {
                        method: "POST",
                        mode: "no-cors",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ action: "delete_region", sap_region_code: regCode })
                    });
                } catch (e) {
                    console.warn(e);
                }
            }

            renderAdminKpisAndSummaries();
            renderAdminZoneProgress();
            renderAdminProductionMatrix();
            renderAdminRegionsTable(document.getElementById('admin-region-search')?.value || '');
            showToast(`🗑️ Cleared data for ${r.region_name}.`);
        }

        async function deleteAllCampaignData() {
            const promptVal = prompt("🚨 DANGER ZONE: Type 'DELETE ALL' to clear all 1,856 territories across Bangladesh from portal and Google Sheet:");
            if (promptVal !== 'DELETE ALL') {
                showToast("ℹ️ Deletion canceled.");
                return;
            }

            store = {};
            localStorage.removeItem('EXIUM_SWEATER_STORE');

            const url = (cloudApiUrl && cloudApiUrl.startsWith('http')) ? cloudApiUrl : DEFAULT_CLOUD_URL;
            if (url) {
                try {
                    showToast("🔄 Clearing Google Sheet...");
                    await fetch(url, {
                        method: "POST",
                        mode: "no-cors",
                        headers: { "Content-Type": "application/json" },
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
            showToast("🗑️ All campaign data has been cleared!");
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

        function exportSingleRegionExcel() {
            onDataChanged();
            const r = REGION_MAP[currentRegionCode];
            if (!r) return;
            const filename = `Exium_Sweater_${r.region_name.replace(/\\s+/g, '_')}_Export.xlsx`;
            generateAndDownloadExcel(r.territories, filename);
        }

        function generateAndDownloadExcel(territoryList, filename) {
            const currentStore = Object.assign({}, JSON.parse(localStorage.getItem('EXIUM_SWEATER_STORE') || '{}'), store);

            const c1Rows = [];
            const c2Rows = [];

            territoryList.forEach(t => {
                const terrCode = String(t['SAP Territory Code']).trim();
                const d = currentStore[terrCode] || {};

                // Campaign 1 Row
                const c1Ok = d.c1_doc_name && d.c1_doc_rpl && String(d.c1_doc_rpl).length === 6 &&
                             d.c1_m1_sweater && d.c1_m1_size &&
                             d.c1_m2_sweater && d.c1_m2_size &&
                             d.c1_m3_sweater && d.c1_m3_size &&
                             (!d.c1_m4_sweater || (d.c1_m4_sweater && d.c1_m4_size));

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
                    "Territory Status": c1Ok ? "Complete" : ((d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater) ? "In Progress" : "Not Started")
                });

                // Campaign 2 Row (3 Doctors)
                const c2Ok = d.c2_d1_name && d.c2_d1_rpl && String(d.c2_d1_rpl).length === 6 && d.c2_d1_sweater && d.c2_d1_size &&
                             d.c2_d2_name && d.c2_d2_rpl && String(d.c2_d2_rpl).length === 6 && d.c2_d2_sweater && d.c2_d2_size &&
                             d.c2_d3_name && d.c2_d3_rpl && String(d.c2_d3_rpl).length === 6 && d.c2_d3_sweater && d.c2_d3_size;

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
                    "Territory Status": c2Ok ? "Complete" : ((d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater || d.c2_d2_name || d.c2_d3_name) ? "In Progress" : "Not Started")
                });
            });

            const wb = XLSX.utils.book_new();
            const ws1 = XLSX.utils.json_to_sheet(c1Rows);
            const ws2 = XLSX.utils.json_to_sheet(c2Rows);

            XLSX.utils.book_append_sheet(wb, ws1, "Gyne Core Doctor (Family)");
            XLSX.utils.book_append_sheet(wb, ws2, "Core Doctor Maximization");

            XLSX.writeFile(wb, filename);
            showToast("📥 Master Excel downloaded successfully!");
        }

        function openCatalogModal() {
            document.getElementById('catalog-modal').classList.remove('hidden');
        }

        function closeCatalogModal() {
            document.getElementById('catalog-modal').classList.add('hidden');
        }

        function openImageLightbox(code) {
            const details = SWEATER_DETAILS[code];
            if (!details) return;

            document.getElementById('lightbox-title').textContent = `${details.code} - ${details.name}`;
            document.getElementById('lightbox-subtitle').textContent = `${details.gender} • ${details.color}`;
            document.getElementById('lightbox-img').src = details.img;

            document.getElementById('lightbox-details').innerHTML = `
                <div class="flex justify-between border-b border-slate-700 pb-1">
                    <span class="text-slate-400">Available Sizes:</span>
                    <strong class="text-orange-400 font-mono">${details.sizes.join(', ')}</strong>
                </div>
                <div class="flex justify-between pt-0.5">
                    <span class="text-slate-400">Supplier:</span>
                    <strong class="text-white">${details.supplier}</strong>
                </div>
            `;

            document.getElementById('lightbox-modal').classList.remove('hidden');
        }

        function zoomSlotImage(selectId) {
            const val = document.getElementById(selectId)?.value;
            if (!val) {
                openCatalogModal();
                return;
            }
            const code = val.split(' ')[0];
            openImageLightbox(code);
        }

        function closeLightbox() {
            document.getElementById('lightbox-modal').classList.add('hidden');
        }

        function showToast(msg) {
            const toast = document.getElementById('toast');
            const toastMsg = document.getElementById('toast-msg');
            toastMsg.textContent = msg;
            toast.classList.remove('translate-y-20', 'opacity-0');
            setTimeout(() => {
                toast.classList.add('translate-y-20', 'opacity-0');
            }, 3500);
        }
    </script>
</body>
</html>
"""

# Replace placeholders
html_content = html_template
html_content = html_content.replace('###ZONE_OPTIONS###', zone_options_html)
html_content = html_content.replace('###ALL_TERRITORIES_JSON###', territories_json)
html_content = html_content.replace('###REGION_MAP_JSON###', region_map_json)
html_content = html_content.replace('###ZONES_JSON###', zones_json)

for k, b64 in images_b64.items():
    html_content = html_content.replace(f'###B64_{k}###', b64)

with open(r"G:\Exium\2026\4Q'26\Sweater\Sweater_Campaign_Portal.html", "w", encoding="utf-8") as f:
    f.write(html_content)

with open(r"G:\Exium\2026\4Q'26\Sweater\index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print("Successfully regenerated web app with 3+1 C1 sweaters and 3 C2 doctors!")
