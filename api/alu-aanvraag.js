const path = require("path");
const fs = require("fs");

module.exports = async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  try {
    // Check if xlsx is available
    let XLSX;
    try {
      XLSX = require("xlsx");
    } catch(e) {
      return res.status(500).json({ error: "xlsx not found", detail: e.message });
    }

    // Check if Excel file exists
    const templatePath = path.join(process.cwd(), "alu_aanvraag_excel_sheet_xltx.xlsx");
    if (!fs.existsSync(templatePath)) {
      // List files in root to debug
      const files = fs.readdirSync(process.cwd());
      return res.status(500).json({ error: "Excel file not found", path: templatePath, files: files });
    }

    const d = req.body;
    const templateBuffer = fs.readFileSync(templatePath);
    const wb = XLSX.read(templateBuffer, { type: "buffer" });
    const ws = wb.Sheets["Invulblad"];

    function set(coord, value) {
      if (!value || value === "") return;
      if (!ws[coord]) ws[coord] = { t: "s" };
      ws[coord].v = value;
      ws[coord].w = value;
      ws[coord].t = "s";
    }

    set("B2", d.projectnummer);
    set("B3", d.projectomschrijving);
    set("B4", d.relatie);
    set("B5", d.aanvraagdatum);
    set("B6", d.indiendatum);
    set("B8", d.kleurafwerking);
    set("B10", d.kleur_kaders_buiten);
    set("C10", d.kleur_kaders_binnen);
    set("B11", d.kleur_draai_buiten);
    set("C11", d.kleur_draai_binnen);
    set("B13", d.gevel_serie);
    set("B14", d.gevel_isolatieklasse);
    set("B15", d.voordeurpaneel);
    set("B16", d.onderdorpel);
    set("B18", d.kleur_deurkrukken);
    set("B19", d.kleur_raamkrukken);
    set("B20", d.kleur_scharnieren);
    set("B21", d.balkondeuren);
    set("B22", d.deurkruk);
    set("B23", d.duwer);
    set("B24", d.cilinder_type_gevel);
    set("C24", d.cilinder_merk_gevel ? "T.b.v merk " + d.cilinder_merk_gevel : "T.b.v merk …");
    set("B25", d.cilinder_binnenbuitengevel);
    set("C25", d.cilinder_binnenbuitenmerk_gevel ? "T.b.v merk " + d.cilinder_binnenbuitenmerk_gevel : "T.b.v merk …");
    set("B26", d.knop_cilinder_gevel || "T.b.v merk …");
    set("B28", d.schuif_serie);
    set("B29", d.afdekkap_boven);
    set("B30", d.afdekkap_beneden);
    set("B31", d.cilinder_type_schuif);
    set("C31", d.cilinder_merk_schuif ? "T.b.v merk " + d.cilinder_merk_schuif : "T.b.v merk …");
    set("B33", d.kleur_deurgreep);
    set("B34", d.deurgreep_komgreep);
    set("B35", d.cilinder_schuif);
    set("B36", d.cilinder_type_schuif);
    set("C36", d.cilinder_merk_schuif ? "T.b.v merk " + d.cilinder_merk_schuif : "T.b.v merk …");
    set("B37", d.knop_cilinder_schuif || "T.b.v merk …");
    set("B39", d.soort_beglazing);
    set("B40", d.warm_edge);
    set("B41", d.zonwerend_merk || "…");
    set("B42", d.zonwerend_waarde || "...");
    set("B43", d.geluidwerend_merk || "…");
    set("B44", d.geluidwerend_waarde || "…");
    set("B45", d.nen3569_merk || "…");
    set("B46", d.doorvalveilig_merk || "…");
    set("B47", d.brandwerend_merk || "…");
    set("B48", d.buitenbeglazing_merk || "…");
    set("B49", d.paneel);
    set("B54", d.windgebied);
    set("B55", d.bebouwd);
    set("B56", d.gebouwhoogte || "…");

    const outBuffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=\"alu_aanvraag.xlsx\"");
    res.status(200).send(outBuffer);

  } catch(e) {
    return res.status(500).json({ error: e.message, stack: e.stack });
  }
};
