import FileSaver from 'file-saver'
const Excel = require('exceljs')

var setMerge = function (luckyMerge = {}, worksheet) {
    const mergearr = Object.values(luckyMerge)
    mergearr.forEach(function (elem) {
        // elem格式：{r: 0, c: 0, rs: 1, cs: 2}
        // 按开始行，开始列，结束行，结束列合并（相当于 K10:M12）
        worksheet.mergeCells(
            elem.r + 1,
            elem.c + 1,
            elem.r + elem.rs,
            elem.c + elem.cs
        )
    })
}

var setBorder = function (luckyBorderInfo, worksheet) {
    if (!Array.isArray(luckyBorderInfo)) return
    // console.log('luckyBorderInfo', luckyBorderInfo)
    luckyBorderInfo.forEach(function (elem) {
        // 现在只兼容到borderType 为range的情况
        // console.log('ele', elem)
        if (elem.rangeType === 'range') {
            const border = borderConvert(elem.borderType, elem.style, elem.color)
            const rang = elem.range[0]
            // console.log('range', rang)
            const row = rang.row
            const column = rang.column
            for (let i = row[0] + 1; i < row[1] + 2; i++) {
                for (let y = column[0] + 1; y < column[1] + 2; y++) {
                    worksheet.getCell(i, y).border = border
                }
            }
        }
        if (elem.rangeType === 'cell') {
            // col_index: 2
            // row_index: 1
            // b: {
            //   color: '#d0d4e3'
            //   style: 1
            // }
            const { colIndex, rowIndex } = elem.value
            const borderData = Object.assign({}, elem.value)
            delete borderData.col_index
            delete borderData.row_index
            const border = addborderToCell(borderData, rowIndex, colIndex)
            // console.log('bordre', border, borderData)
            worksheet.getCell(rowIndex + 1, colIndex + 1).border = border
        }
        // console.log(rang.column_focus + 1, rang.row_focus + 1)
        // worksheet.getCell(rang.row_focus + 1, rang.column_focus + 1).border = border
    })
}
var setStyleAndValue = function (cellArr, worksheet) {
    if (!Array.isArray(cellArr)) return
    cellArr.forEach(function (row, rowid) {
        row.every(function (cell, columnid) {
            if (!cell) return true
            const fill = fillConvert(cell.bg)

            const font = fontConvert(
                cell.ff,
                cell.fc,
                cell.bl,
                cell.it,
                cell.fs,
                cell.cl,
                cell.ul
            )
            const alignment = alignmentConvert(cell.vt, cell.ht, cell.tb, cell.tr)
            let value = ''

            if (cell.f) {
                value = { formula: cell.f, result: cell.v }
            } else if (!cell.v && cell.ct && cell.ct.s) {
                // xls转为xlsx之后，内部存在不同的格式，都会进到富文本里，即值不存在与cell.v，而是存在于cell.ct.s之后
                // value = cell.ct.s[0].v
                cell.ct.s.forEach(arr => {
                    value += arr.v
                })
            } else {
                value = cell.v
            }
            //  style 填入到_value中可以实现填充色
            const letter = createCellPos(columnid)
            const target = worksheet.getCell(letter + (rowid + 1))
            // console.log('1233', letter + (rowid + 1))
            for (const key in fill) {
                target.fill = fill
                console.log(key)
                break
            }
            target.font = font
            target.alignment = alignment
            target.value = value

            return true
        })
    })
}

var setImages = function (imagesArr, worksheet, workbook) {
    if (typeof imagesArr !== 'object') return
    for (const key in imagesArr) {
        // console.log(imagesArr[key]);
        // 通过 base64  将图像添加到工作簿
        const myBase64Image = imagesArr[key].src
        // 开始行 开始列 结束行 结束列
        const start = { col: imagesArr[key].fromCol, row: imagesArr[key].fromRow }
        const end = { col: imagesArr[key].toCol, row: imagesArr[key].toRow }
        const imageId = workbook.addImage({
            base64: myBase64Image,
            extension: 'png'
        })
        worksheet.addImage(imageId, {
            tl: start,
            br: end,
            editAs: 'oneCell'
        })
    }
}

var fillConvert = function (bg) {
    if (!bg) {
        return {}
    }
    // const bgc = bg.replace('#', '')
    const fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: bg.replace('#', '') }
    }
    return fill
}

var fontConvert = function (
    ff = 0,
    fc = '#000000',
    bl = 0,
    it = 0,
    fs = 10,
    cl = 0,
    ul = 0
) {
    // luckysheet：ff(样式), fc(颜色), bl(粗体), it(斜体), fs(大小), cl(删除线), ul(下划线)
    const luckyToExcel = {
        0: '微软雅黑',
        1: '宋体（Song）',
        2: '黑体（ST Heiti）',
        3: '楷体（ST Kaiti）',
        4: '仿宋（ST FangSong）',
        5: '新宋体（ST Song）',
        6: '华文新魏',
        7: '华文行楷',
        8: '华文隶书',
        9: 'Arial',
        10: 'Times New Roman ',
        11: 'Tahoma ',
        12: 'Verdana',
        num2bl: function (num) {
            return num !== 0
        }
    }
    // 出现Bug，导入的时候ff为luckyToExcel的val

    const font = {
        name: typeof ff === 'number' ? luckyToExcel[ff] : ff,
        family: 1,
        size: fs,
        color: { argb: fc.replace('#', '') },
        bold: luckyToExcel.num2bl(bl),
        italic: luckyToExcel.num2bl(it),
        underline: luckyToExcel.num2bl(ul),
        strike: luckyToExcel.num2bl(cl)
    }

    return font
}

var alignmentConvert = function (
    vt = 'default',
    ht = 'default',
    tb = 'default',
    tr = 'default'
) {
    // luckysheet:vt(垂直), ht(水平), tb(换行), tr(旋转)
    const luckyToExcel = {
        vertical: {
            0: 'middle',
            1: 'top',
            2: 'bottom',
            default: 'top'
        },
        horizontal: {
            0: 'center',
            1: 'left',
            2: 'right',
            default: 'left'
        },
        wrapText: {
            0: false,
            1: false,
            2: true,
            default: false
        },
        textRotation: {
            0: 0,
            1: 45,
            2: -45,
            3: 'vertical',
            4: 90,
            5: -90,
            default: 0
        }
    }

    const alignment = {
        vertical: luckyToExcel.vertical[vt],
        horizontal: luckyToExcel.horizontal[ht],
        wrapText: luckyToExcel.wrapText[tb],
        textRotation: luckyToExcel.textRotation[tr]
    }
    return alignment
}

var borderConvert = function (borderType, style = 1, color = '#000') {
    // 对应luckysheet的config中borderinfo的的参数
    if (!borderType) {
        return {}
    }
    const luckyToExcel = {
        type: {
            'border-all': 'all',
            'border-top': 'top',
            'border-right': 'right',
            'border-bottom': 'bottom',
            'border-left': 'left'
        },
        style: {
            0: 'none',
            1: 'thin',
            2: 'hair',
            3: 'dotted',
            4: 'dashDot', // 'Dashed',
            5: 'dashDot',
            6: 'dashDotDot',
            7: 'double',
            8: 'medium',
            9: 'mediumDashed',
            10: 'mediumDashDot',
            11: 'mediumDashDotDot',
            12: 'slantDashDot',
            13: 'thick'
        }
    }
    const template = {
        style: luckyToExcel.style[style],
        color: { argb: color.replace('#', '') }
    }
    const border = {}
    if (luckyToExcel.type[borderType] === 'all') {
        border['top'] = template
        border['right'] = template
        border['bottom'] = template
        border['left'] = template
    } else {
        border[luckyToExcel.type[borderType]] = template
    }
    // console.log('border', border)
    return border
}

function addborderToCell (borders) {
    const border = {}
    const luckyExcel = {
        type: {
            l: 'left',
            r: 'right',
            b: 'bottom',
            t: 'top'
        },
        style: {
            0: 'none',
            1: 'thin',
            2: 'hair',
            3: 'dotted',
            4: 'dashDot', // 'Dashed',
            5: 'dashDot',
            6: 'dashDotDot',
            7: 'double',
            8: 'medium',
            9: 'mediumDashed',
            10: 'mediumDashDot',
            11: 'mediumDashDotDot',
            12: 'slantDashDot',
            13: 'thick'
        }
    }
    // console.log('borders', borders)
    for (const bor in borders) {
        // console.log(bor)
        if (borders[bor].color.indexOf('rgb') === -1) {
            border[luckyExcel.type[bor]] = {
                style: luckyExcel.style[borders[bor].style],
                color: { argb: borders[bor].color.replace('#', '') }
            }
        } else {
            border[luckyExcel.type[bor]] = {
                style: luckyExcel.style[borders[bor].style],
                color: { argb: borders[bor].color }
            }
        }
    }

    return border
}

function createCellPos (n) {
    const ordA = 'A'.charCodeAt(0)

    const ordZ = 'Z'.charCodeAt(0)
    const len = ordZ - ordA + 1
    let s = ''
    while (n >= 0) {
        s = String.fromCharCode((n % len) + ordA) + s

        n = Math.floor(n / len) - 1
    }
    return s
}
export var exportExcel = function (luckysheet, value) {
    // 1.创建工作簿，可以为工作簿添加属性
    const workbook = new Excel.Workbook()
    // 2.创建表格，第二个参数可以配置创建什么样的工作表
    luckysheet.forEach(function (table) {
        if (table.data.length === 0) return true
        const worksheet = workbook.addWorksheet(table.name)
        const merge = (table.config && table.config.merge) || {}
        const borderInfo = (table.config && table.config.borderInfo) || {}
        // 3.设置单元格合并,设置单元格边框,设置单元格样式,设置值,导出图片
        setStyleAndValue(table.data, worksheet)
        setMerge(merge, worksheet)
        setBorder(borderInfo, worksheet)
        setImages(table.images, worksheet, workbook)
        return true
    })

    // 4.写入 buffer
    const buffer = workbook.xlsx.writeBuffer().then(data => {
        const blob = new Blob([data], {
            type: 'application/vnd.ms-excel;charset=utf-8'
        })
        console.log('导出成功！')
        FileSaver.saveAs(blob, `${value}.xlsx`)
    })
    return buffer
}