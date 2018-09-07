import "stdlib/native"
import * as web from "stdlib/web";

export function getPluginInfo() {
    return {
        name: "@@Navision",
        description: "@@Exportación Microsoft Navision"
    }
}

function init() {
    web.addRoute("/amura/navision/export.xlsx", exportHandler, web.adminFilter);
}

function exportHandler(c: web.Context) {
    let start = c.request.date("start");
    let end = c.request.date("end");

    if (!start) {
        throw "s:" + T("@@Especifique la fecha de inicio")
    }

    if (!end) {
        end = time.now();
    }

    let payments = sql.query(`SELECT p.id,
                                     p.createDate,
                                     p.quantity,
                                     p.amount,
                                     t.name AS taxName,
                                     s.idProduct,
                                     s.discount
                              FROM amura:billing:payment p
                              JOIN amura:billing:saleline s ON s.id = p.idSaleline
                              JOIN amura:billing:taxtype t ON t.id = s.idTax
                              WHERE p.createDate BETWEEN ? AND ?`, start, end);

    let f = xlsx.newFile();

    let sheet = f.addSheet("Header");
    let row = sheet.addRow();
    row.addCell("Fecha inicio");
    row.addCell(start.format("02/01/2006")); // For format options (https://golang.org/src/time/format.go)
    row = sheet.addRow();
    row.addCell("Fecha fin");
    row.addCell(end.format("02/01/2006"));

    sheet = f.addSheet("Lines");
    row = sheet.addRow();
    row.addCell("");
    row.addCell("Nº documento");
    row.addCell("Nº línea");
    row.addCell("Tipo");
    row.addCell("Nº");
    row.addCell("Cantidad");
    row.addCell("Precio venta");
    row.addCell("% Descuento línea");
    row.addCell("Factura-Nº cliente");
    row.addCell("Grupo registro IVA prod.");
    row.addCell("Nº obra");
    row.addCell("Agrupación coste");
    for (let p of payments) {
        row = sheet.addRow();
        row.addCell("Factura");
        row.addCell(p.createDate.format("06/01"));
        row.addCell(p.id);
        row.addCell("Producto");
        row.addCell(p.idProduct);
        row.addCell(p.quantity);
        row.addCell(p.amount);
        row.addCell(p.discount / 100);
        row.addCell(43010009);
        row.addCell("GASTO " + p.taxName);
        row.addCell("02");
        row.addCell("01");
    }

    let r = c.response;
    r.setHeader("Content-Disposition", "attachment; filename=data.xlsx");
    r.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    f.write(r);
}
