const path = require("path");
const fs = require("fs");

module.exports = async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();

  try {
    const ExcelJS = require("exceljs");
    const d = req.body;

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Invulblad");

    // Column widths
    ws.getColumn('A').width = 44;
    ws.getColumn('B').width = 26;
    ws.getColumn('C').width = 22;

    // Styles
    const thinBorder = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };

    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    const greyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } };
    const lightGreyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

    function styleCell(cell, opts = {}) {
      cell.border = thinBorder;
      if (opts.fill) cell.fill = opts.fill;
      if (opts.bold) cell.font = { name: 'Arial', bold: true, size: opts.size || 11, color: opts.color ? { argb: opts.color } : undefined };
      else cell.font = { name: 'Arial', size: opts.size || 11 };
      cell.alignment = { vertical: 'middle', horizontal: opts.align || 'left', wrapText: opts.wrap !== false };
    }

    function val(v) { return v && v !== "Niet van toepassing" && v !== "maak keuze" ? v : null; }

    // Row 1 - Title
    ws.getRow(1).height = 24;
    ws.mergeCells('A1:C1');
    const r1 = ws.getCell('A1');
    r1.value = 'Aanvraag formulier ALUMINIUM';
    styleCell(r1, { fill: headerFill, bold: true, size: 13, color: 'FFFFFFFF', align: 'center', wrap: false });

    // Rows 2-6 project info
    const projectRows = [
      [2, 'Projectnummer:', d.projectnummer],
      [3, 'Projectomschrijving:', d.projectomschrijving],
      [4, 'Relatie:', d.relatie],
      [5, 'Aanvraagdatum:', d.aanvraagdatum],
      [6, 'Indiendatum:', d.indiendatum],
    ];
    projectRows.forEach(([row, label, value]) => {
      ws.getRow(row).height = 16;
      ws.mergeCells(`B${row}:C${row}`);
      const a = ws.getCell(`A${row}`);
      const b = ws.getCell(`B${row}`);
      a.value = label;
      b.value = value || '…';
      styleCell(a);
      styleCell(b);
    });

    // Row 7 empty
    ws.getRow(7).height = 16;
    ws.mergeCells('B7:C7');
    styleCell(ws.getCell('A7'));
    styleCell(ws.getCell('B7'));

    // Row 8 Kleurafwerking
    ws.getRow(8).height = 18;
    ws.mergeCells('B8:C8');
    const a8 = ws.getCell('A8');
    const b8 = ws.getCell('B8');
    a8.value = 'Kleurafwerking:';
    b8.value = val(d.kleurafwerking) || 'Niet van toepassing';
    styleCell(a8, { fill: greyFill, bold: true });
    styleCell(b8);

    // Row 9 headers
    ws.getRow(9).height = 16;
    const b9 = ws.getCell('B9');
    const c9 = ws.getCell('C9');
    b9.value = 'Kleur buitenzijde';
    c9.value = 'Kleur binnenzijde';
    styleCell(b9, { fill: lightGreyFill, bold: true, size: 10, align: 'center', wrap: false });
    styleCell(c9, { fill: lightGreyFill, bold: true, size: 10, align: 'center', wrap: false });

    // Row 10-11 kleuren
    [[10, 'Kaders:', d.kleur_kaders_buiten, d.kleur_kaders_binnen],
     [11, 'Draai-/schuifdelen:', d.kleur_draai_buiten, d.kleur_draai_binnen]].forEach(([row, label, vb, vc]) => {
      ws.getRow(row).height = 16;
      styleCell(ws.getCell(`A${row}`), { });
      ws.getCell(`A${row}`).value = label;
      ws.getCell(`B${row}`).value = val(vb) || 'Niet van toepassing';
      ws.getCell(`C${row}`).value = val(vc) || 'Niet van toepassing';
      styleCell(ws.getCell(`B${row}`));
      styleCell(ws.getCell(`C${row}`));
    });

    // Helper for section header
    function sectionHeader(row, text) {
      ws.getRow(row).height = 18;
      ws.mergeCells(`A${row}:C${row}`);
      const cell = ws.getCell(`A${row}`);
      cell.value = text;
      styleCell(cell, { fill: greyFill, bold: true, wrap: false });
    }

    // Helper for merged data row B:C
    function dataRowBC(row, label, value) {
      ws.getRow(row).height = 16;
      ws.mergeCells(`B${row}:C${row}`);
      ws.getCell(`A${row}`).value = label;
      ws.getCell(`B${row}`).value = val(value) || 'Niet van toepassing';
      styleCell(ws.getCell(`A${row}`));
      styleCell(ws.getCell(`B${row}`));
    }

    // Helper for row with separate B and C
    function dataRowBandC(row, label, valueB, valueC) {
      ws.getRow(row).height = 16;
      ws.getCell(`A${row}`).value = label;
      ws.getCell(`B${row}`).value = val(valueB) || 'Niet van toepassing';
      ws.getCell(`C${row}`).value = valueC || 'T.b.v merk …';
      styleCell(ws.getCell(`A${row}`));
      styleCell(ws.getCell(`B${row}`));
      styleCell(ws.getCell(`C${row}`));
    }

    // GEVELKOZIJNEN
    sectionHeader(12, 'GEVELKOZIJNEN:');
    dataRowBC(13, 'Serie:', d.gevel_serie);
    dataRowBC(14, 'Isolatieklasse:', d.gevel_isolatieklasse);
    dataRowBC(15, 'Voordeurpaneel:', d.voordeurpaneel);
    dataRowBC(16, 'Onderdorpel deur', d.onderdorpel);
    sectionHeader(17, 'Bediening ramen en deuren');
    dataRowBC(18, 'Kleur deurkrukken:', d.kleur_deurkrukken);
    dataRowBC(19, 'Kleur raamkrukken:', d.kleur_raamkrukken);
    dataRowBC(20, 'Kleur scharnieren:', d.kleur_scharnieren);
    dataRowBC(21, 'Balkondeuren (kader rondom):', d.balkondeuren);
    dataRowBC(22, 'Deurkruk: ', d.deurkruk);
    dataRowBandC(23, 'Duwer t.b.v voordeur', d.duwer, d.duwer ? null : '…');
    dataRowBandC(24, 'Cilinder type:', d.cilinder_type_gevel, d.cilinder_merk_gevel ? 'T.b.v merk ' + d.cilinder_merk_gevel : 'T.b.v merk …');
    dataRowBandC(25, 'Cilinder binnen /buiten etc.:', d.cilinder_binnenbuitengevel, d.cilinder_binnenbuitenmerk_gevel ? 'T.b.v merk ' + d.cilinder_binnenbuitenmerk_gevel : 'T.b.v merk …');
    dataRowBC(26, 'Knop cilinder icm inbraakwerend glas:', d.knop_cilinder_gevel || 'T.b.v merk …');

    // SCHUIFPUIEN
    sectionHeader(27, 'SCHUIFPUIEN');
    dataRowBC(28, 'Serie:', d.schuif_serie);
    dataRowBC(29, 'Afdekkap 4110 bovenzijde:', d.afdekkap_boven);
    dataRowBC(30, 'Afdekkap 4110 benedenzijde:', d.afdekkap_beneden);
    dataRowBandC(31, 'Cilinder type:', d.cilinder_type_schuif, d.cilinder_merk_schuif ? 'T.b.v merk ' + d.cilinder_merk_schuif : 'T.b.v merk …');
    sectionHeader(32, 'Bediening ');
    dataRowBC(33, 'Kleur deurgreep:', d.kleur_deurgreep);
    dataRowBC(34, 'Deurgreep/komgreep:', d.deurgreep_komgreep);
    dataRowBC(35, 'Cilinder:', d.cilinder_schuif);
    dataRowBandC(36, 'Cilinder type:', d.cilinder_type_schuif, d.cilinder_merk_schuif ? 'T.b.v merk ' + d.cilinder_merk_schuif : 'T.b.v merk …');
    dataRowBC(37, 'Knop cilinder i.c.m inbraakwerend glas :', d.knop_cilinder_schuif || 'T.b.v merk …');

    // BEGLAZING
    sectionHeader(38, 'BEGLAZING');
    dataRowBC(39, 'Soort beglazing:', d.soort_beglazing);
    dataRowBC(40, 'Aanvulling beglazing (warm edge):', d.warm_edge);
    dataRowBC(41, 'Zonwerend t.b.v. merk:', d.zonwerend_merk || '…');
    dataRowBC(42, ' - indien van toepassing zonwerende waarde:', d.zonwerend_waarde || '...');
    dataRowBC(43, 'Geluidwerend t.b.v. merk:', d.geluidwerend_merk || '…');
    dataRowBC(44, ' - indien van toepassing geluidwerende waarde:', d.geluidwerend_waarde || '…');
    dataRowBC(45, 'NEN3569 t.b.v. merk:', d.nen3569_merk || '…');
    dataRowBC(46, 'Doorvalveilig t.b.v. merk: ', d.doorvalveilig_merk || '…');
    dataRowBC(47, 'Brandwerend 30 min. WBDBO t.b.v. merk:', d.brandwerend_merk || '…');
    dataRowBC(48, 'Buitenbeglazing t.b.v. merk:', d.buitenbeglazing_merk || '…');

    // PANEEL
    ws.getRow(49).height = 18;
    ws.mergeCells('B49:C49');
    ws.getCell('A49').value = 'PANEEL:';
    ws.getCell('B49').value = val(d.paneel) || 'Niet van toepassing';
    styleCell(ws.getCell('A49'), { fill: greyFill, bold: true });
    styleCell(ws.getCell('B49'));

    // LET OP row
    ws.getRow(50).height = 16;
    ws.mergeCells('A50:C50');
    ws.getCell('A50').value = 'LET OP: aangeven op tekening in welke kozijnen panelen komen';
    ws.getCell('A50').font = { name: 'Arial', size: 10, color: { argb: 'FFFF0000' } };
    ws.getCell('A50').border = thinBorder;
    ws.getCell('A50').alignment = { vertical: 'middle', wrapText: true };

    // SITUATIE
    sectionHeader(51, 'SITUATIE');
    dataRowBC(52, 'Aanzicht gevels:', 'bijvoegen!');
    dataRowBC(53, 'Plattegronden:', 'indien aanwezig bijvoegen!');
    dataRowBC(54, 'Windgebied:', d.windgebied);
    dataRowBC(55, 'Bebouwd / Onbebouwd:', d.bebouwd);
    dataRowBC(56, 'Gebouwhoogte:', d.gebouwhoogte || '…');

    // Rows 57-63 empty opmerkingen area
    ws.mergeCells('A57:C63');
    for (let r = 57; r <= 63; r++) ws.getRow(r).height = 16;
    styleCell(ws.getCell('A57'), { wrap: true });

    const buf = await wb.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=\"alu_aanvraag.xlsx\"");
    res.status(200).send(buf);

  } catch(e) {
    return res.status(500).json({ error: e.message, stack: e.stack });
  }
};
