/**
 * author: 邹琪珺
 */

let room_num, item_num, investor_list;
function submit() {
    room_num = document.getElementById("room_num").value;
    item_num=document.getElementById("item_num").value;
    run();
}

function selectFile() {
    document.getElementById('file').click();
}

function readWorkbookFromLocalFile(file, callback) {
    var reader = new FileReader();
    reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {type: 'binary'});
        if(callback) callback(workbook);
    };
    reader.readAsBinaryString(file);
}

$(function() {
    document.getElementById('file').addEventListener('change', function(e) {
        var files = e.target.files;
        if(files.length == 0) return;
        var f = files[0];
        if(!/\.xlsx$/g.test(f.name)) {
            alert('仅支持读取xlsx格式！');
            return;
        }
        readWorkbookFromLocalFile(f, function(workbook) {
            let name = workbook.SheetNames[0];
            let aoa = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
            for(let i = 0; i < aoa.length; i++) {
                if(aoa[i].interestString) {
                    if(typeof(aoa[i].interestString) === 'number') {
                        aoa[i].interest = [aoa[i].interestString];
                    } else {
                        let interest = aoa[i].interestString.split(',');
                        aoa[i].interest = interest.map(Number);
                    }
                } else {
                    aoa[i].interest = [];
                }
                aoa[i].order = false;
            }
            investor_list = [...aoa];
            console.log('investor_list', investor_list);
        });
    });
});

// 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
function sheet2blob(sheet, sheetName) {
    sheetName = sheetName || 'sheet1';
    var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
    };
    workbook.Sheets[sheetName] = sheet;
    // 生成excel的配置项
    var wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary'
    };
    var wbout = XLSX.write(workbook, wopts);
    var blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
    // 字符串转ArrayBuffer
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    return blob;
}

/**
 * 通用的打开下载对话框方法，没有测试过具体兼容性
 * @param url 下载地址，也可以是一个blob对象，必选
 * @param saveName 保存文件名，可选
 */
function openDownloadDialog(url, saveName)
{
    if(typeof url == 'object' && url instanceof Blob)
    {
        url = URL.createObjectURL(url); // 创建blob地址
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    var event;
    if(window.MouseEvent) event = new MouseEvent('click');
    else
    {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
}

function exportSpecialExcel(aoa) {
    var sheet = XLSX.utils.aoa_to_sheet(aoa);
    openDownloadDialog(sheet2blob(sheet), '排表.xlsx');
}

// 检查第几个数组是最大值
function maxArr(table) {
    let max = 0;
    let maxIdx = 0;
    for(let i = 0; i < table.length; i++) {
        if(table[i].length > max) {
            max = table[i].length;
            maxIdx = i;
        }
    }
    return maxIdx
}
// 判断是否排完
function checkMustFinish(table) {
    for(let i = 0; i < table.length; i++){
        if(table[i].length !== 0) {
            return false;
        }
    }
    return true;
}
// 判断每列数组的最大值
function maxCol(table) {
    let count = [0, 0, 0, 0, 0, 0, 0, 0];
    for(let i = 0; i < table.length; i++) {
        for(let j = 0; j < table[0].length; j++) {
            if(table[i][j] != 0) {
                count[j] ++;
            }
        }
    }
    let max = Math.max(...count);
    let indexOfMax = count.indexOf(max);
}
// 判断每列是否排满
function judgeFull(table, col) {
    let blank = item_num - room_num;
    let count = 0;
    for(let i = 0; i < item_num; i++) {
        if(table[i][col] === 0) {
            count ++;
        }
    }
    if(count > blank) {
        return false;
    }
    return true;
}

function run() {
    /**
     * 投资人房间安排
     */
    if(room_num > investor_list.length) {
        alert('投资人数量小于房间数量，请重新填写');
        return;
    }

    let order_room = new Array();
    for(let i = 0; i < room_num; i++) {
        order_room[i] = new Array();
    }
    let full = 0;
    let left_investor = investor_list.length;
    // 首先把人数较多的投资人放到单独的房间
    for(let index in investor_list) {
        if(investor_list[index].number >= 3) {
            order_room[full].push(investor_list[index]);
            full++;
            investor_list[index].order = true;
            left_investor--;
        }
    }
    // 把选了超过8个项目的投资人单独放在一间
    for(let index in investor_list) {
        if(investor_list[index].interest.length >= 8 && investor_list[index].order === false) {
            order_room[full].push(investor_list[index]);
            full++;
            investor_list[index].order = true;
            left_investor--;
        }
    }

    // 房间数量小于投资人数量，需要把两组投资人放在一起
    let num1, num2;
    while(room_num - full < left_investor) {
        num1 = -1;
        num2 = -1;
        for(let i = 0; i < investor_list.length; i++) {
            // 查询只上午的投资人
            if(investor_list[i].order === false && investor_list[i].time === 2) {
                num1 = i;
                break;
            }
        }
        // 只有全天的投资人
        if(num1 < 0) {
            for(let i = 0; i < investor_list.length; i++) {
                if(investor_list[i].order === false) {
                    num1 = i;
                    console.log(num1);
                    break;
                }
            }
            for(let i = num1 + 1; i < investor_list.length; i++) {
                if(investor_list[i].order === false) {
                    num2 = i;
                    console.log(num2);
                    break;
                }
            }
            if(num1 < 0 || num2 < 0 || num1 >= investor_list.length || num2 >= investor_list.length) {
                // 有问题
                throw Error('问题1')
            } else {
                order_room[full].push(investor_list[num1]);
                order_room[full].push(investor_list[num2]);
                full++;
                investor_list[num1].order = true;
                investor_list[num2].order = true;
                left_investor = left_investor - 2;
            }
        } else {
            // 存在只上午有空的投资人
            for(let i = 0; i < investor_list.length; i++) {
                if(investor_list[i].order === false && investor_list[i].time === 1) {
                    num2 = i;
                    break;
                }
            }
            if(num2 < 0 || num2 >= investor_list.length) {
                // 有问题
                throw Error('问题2')
            } else {
                order_room[full].push(investor_list[num1]);
                order_room[full].push(investor_list[num2]);
                full++;
                investor_list[num1].order = true;
                investor_list[num2].order = true;
                left_investor = left_investor - 2;
            }
        }
    }

    // 此时，剩余没排的投资人数量跟剩余房间数一致，直接一对一安排房间
    for(let i = 0; i < investor_list.length; i++) {
        if(investor_list[i].order === false) {
            order_room[full].push(investor_list[i])
            full++;
            left_investor--;
            investor_list[i].order = true;
        }
    }

    // 标识只有上午投资人的房间
    for(let i = 0; i < order_room.length; i++) {
        order_room[i].onlyMorning = true;
        for(let j = 0; j < order_room[i].length; j++) {
            if(order_room[i][j].time === 1) {
                order_room[i].onlyMorning = false;
                break;
            }
        }
    }

    // 只打印房间和对应的机构名称
    let room_show = new Array;
    for(let i = 0; i < order_room.length; i++) {
        room_show[i] = new Array;
        for(let j = 0; j < order_room[i].length; j++) {
            room_show[i].push(order_room[i][j].name);
        }
    }

    console.log('room_show', room_show);

    /**
     * 创业者房间安排（生成指定表格）
     */
    let order_interest = new Array();
    for(let i = 0; i < room_num; i++) {
        order_interest[i] = new Array();
    }
    for(let i = 0; i < room_num; i++) {
        if(order_room[i].length === 1) {
            order_interest[i] = order_room[i][0].interest;
        } else {
            
            let arr1 = order_room[i][0].interest;
            let arr2 = order_room[i][1].interest;
            //合并两个数组
            arr1.push(...arr2);//或者arr1 = [...arr1,...arr2]
            //去重
            order_interest[i] = Array.from(new Set(arr1));//let arr3 = [...new Set(arr1)]
        }
    }

    console.log(order_interest);

    // 检查是否重复：表格-数字-行-列
    let order_int = [...order_interest];    // 二维数组的复制
    function checkRepeat(table, num, row, column) {
        let col = new Array()
        for(let i = 0; i < item_num; i++) {
            col.push(table[i][column]);
        }
        return col.includes(num) || table[row].includes(num)
    }


    let temp = [0, 0, 0, 0, 0, 0, 0, 0];
    let order_table = new Array();
    for(let i = 0; i < item_num; i++) {
        order_table[i] = [...temp];
    }

    while(!checkMustFinish(order_int)) {
        let maxIdx = maxArr(order_int);
        let col, item;
        for(let i = 0; i < order_int[maxIdx].length; i++) {
            col = 0;
            item = order_int[maxIdx][i];
            while(order_table[item-1][col] !== 0 || checkRepeat(order_table, maxIdx + 1, item - 1, col)) {
                col++;
                if(col > 7) {
                    throw Error('请更换Excel表中投资人的顺序');
                    return;
                }
            }
            order_table[item-1][col] = maxIdx + 1;
        }
        order_int[maxIdx] = [];
    }

    console.log(order_table);

    /**
     * 创业者房间安排（生成完整表格）
     */
    for(let i = 0; i < 8; i++) {
        // 考虑项目的数量 > 房间数量
        let temp_arr = new Array();
        for(let idx = 0; idx <= room_num; idx++) {
            temp_arr.push(0);
        }
        for(let j = 0; j <= room_num; j++) {
            temp_arr[order_table[j][i]] = 1;
        }
        for(let j = 1; j <= room_num; j++) {
            if(order_room[j-1].onlyMorning === true && i > 5) {
                continue;
            }
            if(temp_arr[j] === 0) {
                // for(let k = item_num - 1; k >= 0; k--) {
                //     if(order_table[k][i] === 0 && !checkRepeat(order_table, j, k, i)) {
                //         order_table[k][i] = j;
                //         break;
                //     }
                // }
                let count = 0;
                while(!judgeFull(order_table, i)  && count < 50) {
                    let temp = Math.floor(Math.random()*11);
                    count++;
                    if(order_table[temp][i] === 0 && !checkRepeat(order_table, j, temp, i)) {
                        order_table[temp][i] = j;
                    }
                }
            }
        }
    }
    console.log(order_table);

    // 加#
    for(let i = 0; i < order_table.length; i++) {
        for(let j = 0; j < order_table[0].length; j++) {
            order_table[i][j] = '#' + order_table[i][j];
        }
    }
    exportSpecialExcel(order_table);
}