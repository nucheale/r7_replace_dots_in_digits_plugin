(function (window, undefined) {
    window.Asc.plugin.init = function () {
        main()
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    function replaceDots() {
        const selectedRange = Api.GetSelection()
        const selectedAddress = selectedRange.Address
        // console.log(selectedAddress.endsWith('1048576'))
        if (/^\$[A-Z]+:\$[A-Z]+$/.test(selectedAddress) || selectedAddress.endsWith('1048576')) {
            // console.log('Выбран целый столбец')
            // console.log(selectedRange.Address)
            console.log(Common.UI.alert({
                title: 'Ошибка',
                msg: `Выбран целый столбец: ${selectedRange.Address}`
            }))
        } else {
            selectedRange.ForEach(function (cell) {
                let cellValue = cell.GetValue();
                if (cellValue.length > 0 && Number(cellValue)) {
                    cell.SetValue(cellValue.replace('.', ","));
                    cell.SetNumberFormat('#,##0.00')
                }
            })
        }
    }

    async function main() {
        return new Promise((resolve) => {
            window.Asc.plugin.callCommand(replaceDots, true, true, function (value) {
                resolve(value);
            })
        })
    }

})(window, undefined);