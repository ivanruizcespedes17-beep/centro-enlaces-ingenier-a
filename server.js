const express = require("express");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 3000;

const ARCHIVO_EXCEL = path.join(
    __dirname,
    "base_datos_links_centro_enlaces.xlsx"
);
const CARPETA_BACKUPS = path.join(__dirname, "backups");

if (!fs.existsSync(CARPETA_BACKUPS)) {
    fs.mkdirSync(CARPETA_BACKUPS);
}
function crearBackup() {
    if (!fs.existsSync(ARCHIVO_EXCEL)) return;

    const ahora = new Date();

    const fecha =
        ahora.getFullYear() + "-" +
        String(ahora.getMonth() + 1).padStart(2, "0") + "-" +
        String(ahora.getDate()).padStart(2, "0") + "_" +
        String(ahora.getHours()).padStart(2, "0") +
        String(ahora.getMinutes()).padStart(2, "0") +
        String(ahora.getSeconds()).padStart(2, "0");

    const archivoBackup = path.join(
        CARPETA_BACKUPS,
        `BACKUP_${fecha}.xlsx`
    );

    fs.copyFileSync(ARCHIVO_EXCEL, archivoBackup);

    console.log(`Backup creado: ${archivoBackup}`);
}
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ===============================
// LEER ENLACES DESDE EXCEL
// ===============================
app.get("/api/links", (req, res) => {
    try {
        const workbook = XLSX.readFile(ARCHIVO_EXCEL);
        const nombreHoja = workbook.SheetNames[0];
        const hoja = workbook.Sheets[nombreHoja];

        const datos = XLSX.utils.sheet_to_json(hoja, {
            defval: ""
        });

        const enlaces = datos.map((fila, indice) => ({
            id: indice + 1,
            name: fila["Nombre del enlace"] || "",
            area: fila["Área"] || "",
            url: fila["URL"] || ""
        }));

        res.json(enlaces);

    } catch (error) {
        console.error("Error leyendo Excel:", error);
        res.status(500).json({
            error: "No se pudo leer el archivo Excel",
            detalle: error.message
        });
    }
});
// =============================
// AGREGAR ENLACE AL EXCEL
// =============================
app.post("/api/add", (req, res) => {
    try {
        const { name, area, url } = req.body;

        if (!name || !area || !url) {
            return res.status(400).json({
                ok: false,
                error: "Faltan datos del enlace."
            });
        }

        const workbook = XLSX.readFile(ARCHIVO_EXCEL);
        const nombreHoja = workbook.SheetNames[0];
        const hoja = workbook.Sheets[nombreHoja];

        const datos = XLSX.utils.sheet_to_json(hoja, {
            defval: ""
        });

        datos.push({
            "Nombre del enlace": name,
            "Área": area,
            "URL": url
        });

        const nuevaHoja = XLSX.utils.json_to_sheet(datos);
        workbook.Sheets[nombreHoja] = nuevaHoja;

 	crearBackup();       
	XLSX.writeFile(workbook, ARCHIVO_EXCEL);

        res.json({
            ok: true,
            count: datos.length
        });

    } catch (error) {
        console.error("Error agregando enlace:", error);
        res.status(500).json({
            ok: false,
            error: "No se pudo guardar el enlace en Excel.",
            detalle: error.message
        });
    }
});
// ===============================
// INICIAR SERVIDOR

// ===============================
// ================================
// GUARDAR TODOS LOS ENLACES EN EXCEL
// ================================
app.post("/api/links", (req, res) => {
    try {
        const links = req.body.links;

        if (!Array.isArray(links)) {
            return res.status(400).json({
                ok: false,
                error: "Los datos recibidos no son válidos."
            });
        }

        const workbook = XLSX.readFile(ARCHIVO_EXCEL);
        const nombreHoja = workbook.SheetNames[0];

        const datos = links.map(link => ({
            "Nombre del enlace": link.name || "",
            "Área": link.area || "",
            "URL": link.url || ""
        }));

        const nuevaHoja = XLSX.utils.json_to_sheet(datos);

        workbook.Sheets[nombreHoja] = nuevaHoja;

	crearBackup();
        XLSX.writeFile(workbook, ARCHIVO_EXCEL);

        res.json({
            ok: true,
            count: links.length
        });

    } catch (error) {
        console.error("Error guardando enlaces:", error);

        res.status(500).json({
            ok: false,
            error: "No se pudo guardar en el Excel.",
            detalle: error.message
        });
    }
});
app.listen(PORT, "0.0.0.0", () => {
    console.log("");
    console.log("======================================");
    console.log(" CENTRO DE ENLACES INGENIERÍA");
    console.log("======================================");
    console.log(`Servidor: http://localhost:${PORT}`);
    console.log(`Excel: ${ARCHIVO_EXCEL}`);
    console.log("");
});