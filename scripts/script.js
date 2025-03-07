(function (window, undefined) {
    window.Asc.plugin.init = function () {
        main()
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    function replaceDots() {
        const activeSheet = Api.GetActiveSheet()
        const selectedRange = Api.GetSelection()
        const selectedAddress = selectedRange.Address
        if (/^\$[A-Z]+:\$[A-Z]+$/.test(selectedAddress) || selectedAddress.endsWith('1048576')) { // если выделили столбец полностью
            // console.log('Выбран целый столбец')
            // console.log(selectedRange.Address)
            console.log(Common.UI.alert({
                title: 'Ошибка',
                msg: `Выбран целый столбец: ${selectedRange.Address}`
            }))
            return
        } else {

            if (selectedRange.areas.length > 1) { // если выделели несколько диапазонов через Ctrl
                let ranges = []
                selectedRange.areas.forEach(area => { // находим адреса этих диапазонов
                    let range = []
                    let startRow = area.kb.r1
                    let startColumn = area.kb.ia
                    let endRow = area.kb.r2
                    let endColumn = area.kb.ra
                    if (endRow == 1048575) { // если один из диапазонов - целый столбец
                        console.log(Common.UI.alert({
                            title: 'Ошибка',
                            msg: `Выбран целый столбец`
                        }))
                        return
                    } 
                    range.push(startRow, startColumn, endRow, endColumn)
                    ranges.push(range)
                })
                ranges.forEach(range => {
                    for (let i = range[0]; i <= range[2]; i++) {
                        for (let j = range[1]; j <= range[3]; j++) {
                            let cell = activeSheet.GetRangeByNumber(i, j)
                            replaceDotsInCell(cell)
                        }
                    }
                })
            } else { // если один диапазон
                selectedRange.ForEach(cell => {
                    replaceDotsInCell(cell)
                })
            }
        }

        function replaceDotsInCell(cell) {
            let cellValue = cell.GetValue();
            if (cellValue.length > 0 && Number(cellValue)) {
                cell.SetValue(cellValue.replace('.', ','));
                cell.SetNumberFormat('#,##0.00');
            }
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