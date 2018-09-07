namespace amura_navision {

    hooks.add(amura.HOOK_ENTITYGRID_CREATED, "amura.billing.payment", args => {
        let grid = args as amura.EntityGrid;
        let items = grid.optionsMenu.items;

        // place the menu entry after the first one (Export to Excel)
        let len = items.length;
        let i = len >= 1 ? 1 : 0;

        grid.optionsMenu.items.insertAt(i, {
            label: T("@@Exportar a Navision"),
            icon: "download",
            url: showExportView
        })
    })

    function showExportView() {
        let form = new S.Form();
        form.addProperty({ name: "start", type: "date", label: T("@@Inicio") })
        form.addProperty({ name: "end", type: "date", label: T("@@Fin"), nullable: true });

        let modal = S.buildModalWindow(T("@@Exportar a Navision"), form, "withPadding");
        modal.onAccept = () => {
            if (!form.validate()) {
                return;
            }
            modal.close();
            S.executeAction("URL:/amura/navision/export.xlsx", form.validator.model)
        };

        modal.show();
    }

}