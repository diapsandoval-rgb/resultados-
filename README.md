# resultados-
<html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Sistema de Simulacros - I.E. La Esperanza</title>
    <script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style id="mainStyles">
        :root {
            --primario: #1a237e; --acento: #3d5afe; --fondo: #f8fafc;
            --blanco: #ffffff; --texto: #1e293b;
            --sob: #10b981; --nor: #f59e0b; --baj: #f97316; --cri: #ef4444;
            --nube: #059669;
        }
        body { font-family: 'Inter', 'Segoe UI', sans-serif; background: var(--fondo); color: var(--texto); margin: 0; }
        .auth-screen { position: fixed; inset: 0; background: linear-gradient(135deg, var(--primario), var(--acento)); display: flex; justify-content: center; align-items: center; z-index: 2000; }
        .auth-card { background: var(--blanco); padding: 40px; border-radius: 24px; width: 380px; text-align: center; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1); }
        .auth-card input { width: 100%; padding: 14px; margin: 10px 0; border: 1.5px solid #e2e8f0; border-radius: 12px; font-size: 1rem; box-sizing: border-box; }
        .btn-login { background: var(--acento); color: white; border: none; width: 100%; padding: 14px; border-radius: 12px; cursor: pointer; font-weight: 700; transition: 0.3s; }
        .btn-nube { background: var(--nube); color: white; border: none; padding: 12px 20px; border-radius: 12px; cursor: pointer; font-weight: 700; margin-top: 10px; display: flex; align-items: center; gap: 8px; justify-content: center; }
        .app { padding: 30px; max-width: 1500px; margin: auto; }
        .header { background: var(--blanco); padding: 20px; border-radius: 20px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
        .admin-box { background: #eef2ff; padding: 20px; border-radius: 20px; margin-bottom: 25px; border: 2px dashed var(--acento); display: none; text-align: center; }
        .tabs { display: flex; gap: 15px; overflow-x: auto; padding-bottom: 15px; margin-bottom: 25px; }
        .tab-wrapper { background: var(--blanco); padding: 15px; border-radius: 18px; min-width: 180px; border: 2px solid transparent; box-shadow: 0 4px 6px rgba(0,0,0,0.05); text-align: center; }
        .tab-wrapper.active { border-color: var(--acento); background: #f0f3ff; }
        .btn-view { background: var(--acento); color: white; width: 100%; font-weight: bold; padding: 8px; border-radius: 8px; border: none; cursor: pointer; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin-bottom: 25px; }
        .stat-card { background: var(--blanco); padding: 20px; border-radius: 18px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        .stat-card h5 { margin: 0; color: #64748b; font-size: 0.75rem; }
        .stat-card .val { font-size: 1.6rem; font-weight: 800; color: var(--primario); }
        .charts { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 25px; margin-bottom: 25px; }
        .chart-card { background: var(--blanco); padding: 25px; border-radius: 24px; box-shadow: 0 10px 15px rgba(0,0,0,0.05); height: 400px; }
        table { width: 100%; border-collapse: collapse; background: white; border-radius: 20px; overflow: hidden; }
        th { padding: 15px; background: #f8fafc; color: #94a3b8; font-size: 0.75rem; }
        td { padding: 15px; border-bottom: 1px solid #f1f5f9; text-align: center; }
        .badge { padding: 5px 12px; border-radius: 99px; font-weight: 700; font-size: 0.65rem; }
        .sob { background: #d1fae5; color: #065f46; }
        .nor { background: #fef3c7; color: #92400e; }
        .baj { background: #ffedd5; color: #9a3412; }
        .cri { background: #fee2e2; color: #991b1b; }
        @media print { .no-print { display: none !important; } }
    </style>
</head>
<body>

    <div id="loginPage" class="auth-screen">
        <div class="auth-card">
            <h2 style="color: var(--primario);">I.E. La Esperanza</h2>
            <p>Ingresa tus credenciales para continuar</p>
            <input type="text" id="userName" placeholder="Nombre completo">
            <input type="password" id="userPass" placeholder="Contraseña (1234)">
            <button class="btn-login" onclick="validarIngreso()">Entrar al Sistema</button>
        </div>
    </div>

    <div id="dashboard" class="app" style="display:none;">
        <div class="header">
            <h2 id="welcomeText" style="margin:0;">Panel Académico</h2>
            <button class="no-print" onclick="location.reload()" style="padding:10px 20px; border-radius:10px; cursor:pointer; border:none; background:#f1f5f9;">Cerrar Sesión</button>
        </div>

        <div id="adminPanel" class="admin-box no-print">
            <h3 style="margin-top:0;">Gestión de Datos en la Nube</h3>
            <div style="display: flex; justify-content: center; gap: 30px; align-items: center; flex-wrap: wrap;">
                <div>
                    <p style="margin: 5px 0; font-weight: bold; color: var(--acento);">Paso 1: Cargar Excel</p>
                    <input type="file" id="fileIn" accept=".xlsx">
                </div>
                <div style="border-left: 2px solid #ddd; padding-left: 30px;">
                    <p style="margin: 5px 0; font-weight: bold; color: var(--nube);">Paso 2: Exportar para Nube</p>
                    <button class="btn-nube" onclick="exportarBaseDatosJSON()">
                        <span>📥 Descargar datos.json</span>
                    </button>
                </div>
            </div>
        </div>

        <div id="tabsContainer" class="tabs no-print"></div>

        <div id="mainContent">
            <div id="statsGrid" class="stats-grid"></div>
            <div class="charts">
                <div class="chart-card"><canvas id="pieChart"></canvas></div>
                <div class="chart-card"><canvas id="barChart"></canvas></div>
            </div>
            <div style="background:white; padding:25px; border-radius:24px; box-shadow:0 10px 15px rgba(0,0,0,0.05); overflow-x: auto;">
                <h3 style="margin-top:0;">Ranking de Resultados</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Pos</th><th>Estudiante</th><th>LEC</th><th>MAT</th><th>SOC</th><th>NAT</th><th>ING</th><th>GLOBAL</th><th>NIVEL</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
        </div>
    </div>

<script>
    // Variables de estado
    let db = []; 
    let currentIdx = null;
    let session = null;
    let chartList = [];

    // --- 1. FUNCIÓN DE AUTO-CARGA DESDE LA NUBE ---
    async function cargarDatosDesdeNube() {
        try {
            // Esta función busca el archivo 'datos.json' en la misma carpeta del link
            const respuesta = await fetch('datos.json');
            if (respuesta.ok) {
                const datosServidor = await respuesta.json();
                db = datosServidor;
                localStorage.setItem('ie_esperanza_db', JSON.stringify(db));
                if (db.length > 0) currentIdx = db.length - 1;
                console.log("Sincronización con la nube exitosa.");
            } else {
                // Si no hay archivo en la nube, intenta cargar la copia local del navegador
                db = JSON.parse(localStorage.getItem('ie_esperanza_db')) || [];
                if (db.length > 0) currentIdx = db.length - 1;
            }
        } catch (e) {
            console.log("Modo local: No se detectó archivo en la nube.");
            db = JSON.parse(localStorage.getItem('ie_esperanza_db')) || [];
            if (db.length > 0) currentIdx = db.length - 1;
        }
    }

    // Ejecutar carga al abrir la página
    cargarDatosDesdeNube();

    // --- 2. FUNCIÓN PARA DESCARGAR EL ARCHIVO DE DATOS ---
    function exportarBaseDatosJSON() {
        if (db.length === 0) {
            alert("Sube al menos un archivo Excel antes de exportar.");
            return;
        }

        const dataStr = JSON.stringify(db, null, 2);
        const blob = new Blob([dataStr], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement("a");
        link.href = url;
        link.download = "datos.json";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        alert("✅ Archivo 'datos.json' generado.\n\nSube este archivo a tu repositorio de GitHub Pages para actualizar los datos de todos los dispositivos.");
    }

    // --- LÓGICA DE SESIÓN ---
    function validarIngreso() {
        const user = document.getElementById('userName').value.trim();
        const pass = document.getElementById('userPass').value;

        if(user.toLowerCase() === 'admin' && pass === 'admin123') {
            session = { nombre: "Administrador", rol: "admin" };
        } else if(user !== "" && pass === "1234") {
            session = { nombre: user, rol: "estudiante" };
        } else {
            alert("Credenciales incorrectas."); return;
        }

        document.getElementById('loginPage').style.display = 'none';
        document.getElementById('dashboard').style.display = 'block';
        document.getElementById('welcomeText').innerText = "Hola, " + session.nombre;
        
        if(session.rol === 'admin') document.getElementById('adminPanel').style.display = 'block';
        render();
    }

    // --- PROCESAMIENTO DE EXCEL ---
    document.getElementById('fileIn').addEventListener('change', function(e) {
        const reader = new FileReader();
        reader.onload = (evt) => {
            const wb = XLSX.read(evt.target.result, {type:'array'});
            const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
            
            let procesados = raw.map(r => {
                const g = Math.round((Number(r.LECTURA)+Number(r.MATEMATICAS)+Number(r.SOCIALES)+Number(r.CIENCIAS)+Number(r.INGLES))/5);
                let nv = g >= 80 ? {n:"Sobresaliente",c:"sob"} : g >= 60 ? {n:"Normal",c:"nor"} : g >= 40 ? {n:"Bajo",c:"baj"} : {n:"Crítico",c:"cri"};
                return { nombre: r.NOMBRE, lec: r.LECTURA, mat: r.MATEMATICAS, soc: r.SOCIALES, cie: r.CIENCIAS, ing: r.INGLES, global: g, nivel: nv };
            }).sort((a,b) => b.global - a.global);

            db.push({ nombre: "Simulacro " + (db.length + 1), data: procesados });
            localStorage.setItem('ie_esperanza_db', JSON.stringify(db));
            currentIdx = db.length - 1;
            render();
            alert("¡Excel cargado con éxito! Recuerda descargar el datos.json y subirlo a la nube.");
        };
        reader.readAsArrayBuffer(e.target.files[0]);
    });

    window.setExamen = (i) => { currentIdx = i; render(); };

    function render() {
        const tabs = document.getElementById('tabsContainer');
        tabs.innerHTML = db.map((ex, i) => `
            <div class="tab-wrapper ${i === currentIdx ? 'active' : ''}">
                <div style="font-weight:bold; color:var(--primario)">${ex.nombre}</div>
                <button class="btn-view" onclick="setExamen(${i})">Ver Reporte</button>
            </div>`).join('');

        if(currentIdx !== null) {
            const original = db[currentIdx].data;
            let mostrar = (session.rol === 'estudiante') ? 
                original.filter(s => s.nombre.toUpperCase() === session.nombre.toUpperCase()) : original;
            
            actualizarUI(mostrar, original);
        }
    }

    function actualizarUI(data, original) {
        document.getElementById('tableBody').innerHTML = data.map(s => {
            const puesto = original.findIndex(e => e.nombre === s.nombre) + 1;
            return `<tr>
                <td style="font-weight:bold; color:var(--acento)">#${puesto}</td>
                <td style="text-align:left; font-weight:600;">${s.nombre}</td>
                <td>${s.lec}</td><td>${s.mat}</td><td>${s.soc}</td><td>${s.cie}</td><td>${s.ing}</td>
                <td style="font-weight:bold; color:var(--primario)">${s.global}</td>
                <td><span class="badge ${s.nivel.c}">${s.nivel.n}</span></td>
            </tr>`;
        }).join('');

        const mtr = [{k:'lec', l:'Lectura'}, {k:'mat', l:'Mates'}, {k:'soc', l:'Sociales'}, {k:'cie', l:'Ciencias'}, {k:'ing', l:'Inglés'}, {k:'global', l:'Global'}];
        document.getElementById('statsGrid').innerHTML = mtr.map(m => {
            const av = data.length ? data.reduce((a,b)=>a+b[m.k],0)/data.length : 0;
            return `<div class="stat-card"><h5>${m.l}</h5><div class="val">${Math.round(av)}</div></div>`;
        }).join('');

        chartList.forEach(c => c.destroy());
        chartList = [];
        const lv = {sob:0, nor:0, baj:0, cri:0}; data.forEach(s => lv[s.nivel.c]++);

        chartList.push(new Chart(document.getElementById('pieChart'), {
            type: 'doughnut',
            data: { labels: ['Sobresaliente','Normal','Bajo','Crítico'], datasets: [{data:Object.values(lv), backgroundColor:['#10b981','#f59e0b','#f97316','#ef4444']}] },
            options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
        }));
        
        chartList.push(new Chart(document.getElementById('barChart'), {
            type: 'bar',
            data: { 
                labels: ['LEC','MAT','SOC','CIE','ING'], 
                datasets: [{ label: 'Puntaje Individual', data: ['lec','mat','soc','cie','ing'].map(k => data[0]?.[k] || 0), backgroundColor: '#3d5afe', borderRadius: 8 }] 
            },
            options: { maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } } }
        }));
    }
</script>
</body>
</html>
