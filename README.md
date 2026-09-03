
from IPython.display import display, HTML

# هذا الكود يقوم بتشغيل تطبيقك بالكامل كواجهة ويب داخل Google Colab
app_code = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
    <title>HSA Smart Cost Analyzer - Colab Edition</title>

    <!-- مكتبات قراءة الإكسل والرسومات البيانية -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

    <style>
        :root { --blue: #1565C0; --blue-d: #0D47A1; --bg: #F8F9FB; --card: #FFFFFF; --ink: #263238; --soft: #607D8B; --line: #E3E8EE; --green: #2E7D32; --red: #C62828; --amber: #F57F17; }
        * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
        body { font-family: "Segoe UI", Tahoma, Arial, sans-serif; background: var(--bg); color: var(--ink); direction: rtl; padding-bottom: 80px; }

        header { background: linear-gradient(135deg,var(--blue),var(--blue-d)); color: #fff; padding: 14px 16px; box-shadow: 0 2px 8px rgba(13, 71, 161, 0.25); display: flex; align-items: center; gap: 12px; }
        .logo { background: #fff; color: var(--blue); font-weight: 800; font-size: 20px; border-radius: 12px; padding: 8px 10px; }
        header h1 { font-size: 18px; font-weight: 700; margin:0;} header p { font-size: 12px; opacity: 0.85; margin: 2px 0 0 0; }

        main { max-width: 100%; margin: 0 auto; padding: 14px; }
        .card { background: var(--card); border: 1px solid var(--line); border-radius: 14px; padding: 14px; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05); }
        .card h3 { font-size: 15px; color: var(--blue-d); margin-bottom: 8px; }
        .card p { font-size: 13px; color: #37474F; margin-bottom: 10px; line-height: 1.6;}

        .stat-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin-bottom: 12px; }
        .stat { background: var(--card); border-radius: 14px; padding: 14px; text-align: center; border: 1px solid var(--line); }
        .stat .n { font-size: 22px; font-weight: 800; color: var(--blue); } .stat .l { font-size: 12px; color: var(--soft); margin-top: 4px; }

        .btn { background: var(--blue); color: #fff; border: none; border-radius: 10px; padding: 10px 18px; font-size: 14px; font-weight: 600; width: 100%; cursor: pointer; }
        .btn:hover { background: var(--blue-d); }

        .item-row { display: flex; justify-content: space-between; align-items: center; gap: 8px; border-bottom: 1px dashed var(--line); padding: 10px 2px; }
        .item-row .t { font-size: 13px; font-weight: 600; display: flex; align-items: center; gap: 8px; }
        .pill { background: #FFF8E1; color: #795548; border-radius: 20px; padding: 2px 10px; font-size: 11px; font-weight: 600; }
        .kv { display: flex; justify-content: space-between; font-size: 13px; padding: 8px 0; border-bottom: 1px dotted var(--line); }
        .green { color: var(--green); font-weight:bold; }
        .price-tag { background: #E8F5E9; border-radius: 8px; padding: 10px; font-size: 14px; color: var(--green); font-weight: 700; margin-top: 10px; text-align: center; }
        .form-control { width: 100%; padding: 8px; border-radius: 8px; border: 1px solid var(--line); font-family: inherit; margin-bottom: 10px; }

        #tabbar { position: fixed; bottom: 0; left: 0; right: 0; background: #fff; border-top: 1px solid var(--line); display: flex; z-index: 50; }
        #tabbar button { flex: 1; background: none; border: none; color: var(--soft); font-size: 10px; padding: 7px 2px; display: flex; flex-direction: column; align-items: center; gap: 2px; cursor: pointer; }
        #tabbar button span { font-size: 18px; }
        #tabbar button.active { color: var(--blue); font-weight: 700; border-top: 2px solid var(--blue); }

        .tab-content { display: none; }
        .tab-content.active { display: block; animation: fadeIn 0.3s; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: none; } }

        #toast { position: fixed; bottom: 70px; left: 50%; transform: translateX(-50%); background: #263238; color: #fff; padding: 10px 18px; border-radius: 30px; font-size: 13px; opacity: 0; transition: opacity 0.3s; z-index: 99; white-space: nowrap; pointer-events: none;}
        #toast.show { opacity: 1; }

        .chart-container { position: relative; height: 220px; width: 100%; margin-top: 10px; }
    </style>
</head>
<body>

<header>
  <div class="logo">HSA</div>
  <div>
    <h1>نظام التكاليف الذكي</h1>
    <p>نسخة التطوير - Google Colab</p>
  </div>
</header>

<main id="content">

    <!-- 1. الرئيسية -->
    <div id="tab-home" class="tab-content active">
        <div class="stat-grid">
            <div class="stat"><div class="n" id="stat-items">0</div><div class="l">بند عمل محفوظ</div></div>
            <div class="stat"><div class="n">3</div><div class="l">مادة / مواصفة</div></div>
        </div>
        <div class="card">
            <h3>📂 استيراد شامل للبيانات</h3>
            <p>ارفع ملف إكسل (جداول كميات BOQ أو تحليلات تفصيلية).</p>
            <input type="file" id="excelUpload" accept=".xlsx, .xls" style="display: none;" onchange="processAllExcelSheets(this.files[0])">
            <button class="btn" onclick="document.getElementById('excelUpload').click()">استيراد ملف الإكسل الآن</button>
        </div>
    </div>

    <!-- 2. التكلفة التقديرية -->
    <div id="tab-estimates" class="tab-content">
        <div class="card">
            <h3>📍 البنود المسحوبة لتسعير الموقع</h3>
            <div id="estimate_items_list">
                <p style="text-align:center; color:gray; font-size:12px;">قم باستيراد ملف إكسل لظهور البنود هنا.</p>
            </div>
            <button class="btn" onclick="calculateSiteEstimate()" style="margin-top: 15px;">احسب التكلفة التقديرية (ديناميكياً)</button>
            <div id="site_total_estimate" class="price-tag" style="display:none;">
                إجمالي التكلفة التقديرية: <span id="site_total_value">0</span>
            </div>
        </div>

        <div class="card">
            <h3>📊 تحليل نسب التكلفة</h3>
            <div class="chart-container"><canvas id="estimatesChart"></canvas></div>
        </div>
    </div>

    <!-- 3. المواد والمواصفات -->
    <div id="tab-materials_specs" class="tab-content">
        <div class="card">
            <h3>📈 مؤشر تغير أسعار المواسير (UPVC)</h3>
            <div class="chart-container"><canvas id="materialsChart"></canvas></div>
        </div>
        <div class="card">
            <h3>🧱 مواسير بلاستيك ضغط عالي (UPVC)</h3>
            <label style="font-size:12px; color:gray;">اختر الشركة المصنعة / الماركة:</label>
            <select id="manufacturer_select" class="form-control" onchange="updateMaterialPrice(this)">
                <option value="5000" data-spec="درجة أولى - سميك">أنابيب البلاستيك الأهلية (درجة أولى) - 5,000</option>
                <option value="4200" data-spec="درجة ثانية - متوسط">شركة الأمل (درجة ثانية) - 4,200</option>
                <option value="6500" data-spec="مواصفات ألمانية - ثقيل">أكوا ثيرم (مستورد) - 6,500</option>
            </select>
            <div class="kv"><span>المواصفة:</span><b id="spec_display">درجة أولى - سميك</b></div>
            <div class="kv"><span>السعر المعتمد للنظام:</span><b id="price_display" class="green">5,000</b></div>
        </div>
    </div>

    <!-- 4. عروض الأسعار -->
    <div id="tab-quotations" class="tab-content">
        <div class="card">
            <h3>⚖️ مقارنة عطاءات المقاولين</h3>
            <div class="chart-container"><canvas id="quotationsChart"></canvas></div>
        </div>
    </div>

</main>

<nav id="tabbar">
  <button data-tab="home" class="active"><span>🏠</span>الرئيسية</button>
  <button data-tab="estimates"><span>📊</span>التقديرية</button>
  <button data-tab="materials_specs"><span>🏭</span>المواد</button>
  <button data-tab="quotations"><span>⚖️</span>العروض</button>
</nav>

<div id="toast"></div>

<script>
    // 1. Data Initialization
    const defaultMaterials = { "مرحاض عربي نوع خزف مواصفات اوروبية شامل صندوق الطرد": { price: 12000 }, "مواسير بلاستيك ضغط عالي (UPVC)": { price: 5000 } };
    const defaultResources = { "عمل يد سباكه شبكه صرف صحي": { daily_rate: 15000 }, "عمل يد تكسير وحفر": { daily_rate: 10000 } };

    if(!sessionStorage.getItem('hsa_materials')) sessionStorage.setItem('hsa_materials', JSON.stringify(defaultMaterials));
    if(!sessionStorage.getItem('hsa_resources')) sessionStorage.setItem('hsa_resources', JSON.stringify(defaultResources));

    // Note: Using sessionStorage for Colab so it resets cleanly when re-run.

    // 2. Tab Navigation
    document.querySelectorAll('#tabbar button').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('#tabbar button').forEach(b => b.classList.remove('active'));
            const targetBtn = e.currentTarget;
            targetBtn.classList.add('active');

            document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
            const tabId = 'tab-' + targetBtn.getAttribute('data-tab');
            document.getElementById(tabId).classList.add('active');

            if(tabId === 'tab-estimates') renderEstimateItems();
        });
    });

    function showToast(msg) {
        const toast = document.getElementById('toast');
        toast.textContent = msg;
        toast.classList.add('show');
        setTimeout(() => toast.classList.remove('show'), 3500);
    }

    // 3. Excel Processing
    function processAllExcelSheets(file) {
        if(!file) return;
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            let allExtractedItems = [];

            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                let isDetailedAnalysis = false; let isBOQ = false;

                for (let i = 0; i < Math.min(jsonData.length, 20); i++) {
                    const rowStr = JSON.stringify(jsonData[i] || []);
                    if (rowStr.includes("اولا :المواد") || rowStr.includes("ثانيأ : العما لة")) { isDetailedAnalysis = true; break; }
                    if (rowStr.includes("بيان الاعمال") || rowStr.includes("Description of work")) { isBOQ = true; break; }
                }

                if (isDetailedAnalysis) {
                    let currentItemName = "بند مركب: " + sheetName;
                    let materials = []; let labor = []; let currentSection = null;

                    for (let i = 0; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length === 0) continue;
                        const detail = row[1] ? row[1].toString().trim() : "";
                        const unit = row[3] ? row[3].toString().trim() : "";
                        const qty = parseFloat(row[4]) || 0;
                        const waste = parseFloat(row[5]) || 0;

                        if (detail.includes("الاجمــــــــــا لي") || detail.includes("التفـــــــــــــــــــــــــــا  صيل")) continue;
                        if (!currentSection && detail.length > 20 && !detail.includes("المواصفات الفنية")) currentItemName = detail.substring(0, 50) + "...";
                        if (detail.includes("اولا :المواد")) { currentSection = "materials"; continue; }
                        if (detail.includes("ثانيأ : العما لة")) { currentSection = "labor"; continue; }

                        if (detail && unit && unit !== "الوحد ة" && currentSection) {
                            if (currentSection === "materials") materials.push({name: detail, unit: unit, qty: qty, waste_pct: waste});
                            else labor.push({name: detail, unit: unit, qty: qty, waste_pct: waste});
                        }
                    }
                    if(materials.length > 0 || labor.length > 0) {
                        allExtractedItems.push({ id: 'DET_'+Date.now()+Math.floor(Math.random()*100), itemName: currentItemName, type: "تحليل_تفصيلي", materials: materials, labor: labor });
                    }
                }
                else if (isBOQ) {
                    for (let i = 0; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        if (!row || row.length < 5) continue;
                        const description = row[1] ? row[1].toString().trim() : "";
                        const unit = row[2] ? row[2].toString().trim() : "";
                        const qty = parseFloat(row[3]); const price = parseFloat(row[4]);

                        if (description && !isNaN(qty) && !isNaN(price) && !description.includes("بيان الاعمال")) {
                            allExtractedItems.push({ id: 'BOQ_'+Date.now()+Math.floor(Math.random()*100), itemName: description.substring(0, 50)+"...", unit: unit, qty: qty, directPrice: price, type: "مقطوعية_مباشرة" });
                        }
                    }
                }
            });

            if(allExtractedItems.length > 0) {
                let items = JSON.parse(sessionStorage.getItem('hsa_items')) || [];
                items = items.concat(allExtractedItems);
                sessionStorage.setItem('hsa_items', JSON.stringify(items));
                document.getElementById('stat-items').textContent = items.length;
                showToast(`تم استيراد ${allExtractedItems.length} بند بنجاح.`);
                renderEstimateItems();
            } else {
                showToast("لم يتم العثور على بنود.");
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function renderEstimateItems() {
        const container = document.getElementById('estimate_items_list');
        const items = JSON.parse(sessionStorage.getItem('hsa_items')) || [];
        if(items.length === 0) return;

        container.innerHTML = '';
        items.forEach(item => {
            const pillType = item.type === "تحليل_تفصيلي" ? "تحليل موارد" : "مقطوعية";
            container.innerHTML += `
                <div class="item-row">
                    <span class="t"><input type="checkbox" class="site-item-checkbox" value="${item.id}" checked> ${item.itemName}</span>
                    <span class="pill">${pillType}</span>
                </div>`;
        });
    }

    function updateMaterialPrice(selectElement) {
        const opt = selectElement.options[selectElement.selectedIndex];
        document.getElementById('price_display').textContent = Number(opt.value).toLocaleString();
        document.getElementById('spec_display').textContent = opt.getAttribute('data-spec');

        let mDB = JSON.parse(sessionStorage.getItem('hsa_materials')) || {};
        mDB["مواسير بلاستيك ضغط عالي (UPVC)"] = { price: Number(opt.value) };
        sessionStorage.setItem('hsa_materials', JSON.stringify(mDB));
        showToast('تم تحديث السعر. أعد حساب التكلفة.');
    }

    let estChart = null;
    function updateEstimateChart(m, l, e, p) {
        const ctx = document.getElementById('estimatesChart').getContext('2d');
        if (estChart) { estChart.data.datasets[0].data = [m, l, e, p]; estChart.update(); }
        else {
            estChart = new Chart(ctx, {
                type: 'doughnut',
                data: { labels: ['مواد', 'عمالة/مقطوعية', 'معدات', 'ربح'], datasets: [{ data: [m, l, e, p], backgroundColor: ['#1565C0', '#F57F17', '#607D8B', '#2E7D32'], borderWidth: 0 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
            });
        }
    }

    function calculateSiteEstimate() {
        const mDB = JSON.parse(sessionStorage.getItem('hsa_materials')) || {};
        const rDB = JSON.parse(sessionStorage.getItem('hsa_resources')) || {};
        const items = JSON.parse(sessionStorage.getItem('hsa_items')) || [];

        let tMat = 0, tLab = 0, tEq = 0;
        const checks = document.querySelectorAll('.site-item-checkbox:checked');
        if(checks.length === 0) return showToast('حدد بنداً واحداً على الأقل.');

        checks.forEach(chk => {
            const itm = items.find(i => i.id == chk.value);
            if (!itm) return;
            if (itm.type === "تحليل_تفصيلي") {
                if (itm.materials) itm.materials.forEach(mat => { tMat += (mat.qty * (1 + (mat.waste_pct || 0))) * (mDB[mat.name] ? mDB[mat.name].price : 0); });
                if (itm.labor) itm.labor.forEach(lab => { tLab += (lab.qty * (rDB[lab.name] ? rDB[lab.name].daily_rate : 0)); });
            } else if (itm.type === "مقطوعية_مباشرة") {
                const tot = itm.qty * itm.directPrice;
                tMat += (tot * 0.70); tLab += (tot * 0.30);
            }
        });

        const profit = (tMat + tLab + tEq) * 0.15;
        document.getElementById('site_total_value').textContent = Math.round(tMat + tLab + tEq + profit).toLocaleString();
        document.getElementById('site_total_estimate').style.display = 'block';
        updateEstimateChart(tMat, tLab, tEq, profit);
        showToast("تم الحساب وتحديث الرسم.");
    }

    document.addEventListener('DOMContentLoaded', () => {
        new Chart(document.getElementById('quotationsChart').getContext('2d'), {
            type: 'bar', data: { labels: ['التقديرية', 'مؤسسة البناء', 'الثناء'], datasets: [{ label: 'السعر', data: [14500, 16000, 13800], backgroundColor: ['#1565C0', '#C62828', '#2E7D32'], borderRadius: 4 }] },
            options: { responsive: true, maintainAspectRatio: false }
        });
        new Chart(document.getElementById('materialsChart').getContext('2d'), {
            type: 'line', data: { labels: ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو'], datasets: [{ label: 'السعر', data: [4200, 4250, 4400, 4800, 5000, 5000], borderColor: '#F57F17', backgroundColor: 'rgba(245, 127, 23, 0.1)', fill: true, tension: 0.3 }] },
            options: { responsive: true, maintainAspectRatio: false }
        });
    });
</script>
</body>
</html>
"""

# عرض الكود كنافذة ويب متكاملة (IFrame) داخل Colab
display(HTML(f'<div style="max-width: 450px; height: 750px; margin: auto; border: 1px solid #ccc; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1);"><iframe srcdoc="{app_code.replace('"', "&quot;")}" width="100%" height="100%" style="border:none;"></iframe></div>'))
