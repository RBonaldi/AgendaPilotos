
var dias = setData(15);

var ConjuntoEstilos = [];

var alfabeto = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 
'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

data = [
    ['Hercílio L (FLN-SBFL)'],
    ['Aju/Instrutor NDL-AJU'],
    ['Recife-Guararapes- REC'],
    ['PRL-LTZ'],
];

var listaFuncao = ["CMT", "COP", "CMS"]

var listaNome = ["Daniel Schmitt", "Linda Johnson", "Robert Carroll", "Abigail Velasquez", "Allen Young", "Vincent Rose",
    "Jeremy Scott", "Alexander Vasquez", "James Lewis", "Ray Perez", "Charles Williams", "Jennifer Cabrera",
    "Darren Simpson", "Patricia Torres", "Parker Shaffer", "Carla Hammond", "Jordan Welch", "Pamela Williams", 
    "Michelle Reynolds", "Jennifer Wise"];

var colunas = [{
    title:' ',
    type: 'text', 
    width: 160,
    colspan: '3',
    readOnly:true,
}];

colunas.push({
    title: 'Função',
    type: 'autocomplete',
    source: listaFuncao,
    width: '80px'
});
  
for (var i = 0; i <= dias; i++) {
    var dtColum = new Date();
    dtColum.setDate(dtColum.getDate() + i);
    dtColum = new Intl.DateTimeFormat('pt-BR').format(dtColum);

    colunas.push({
        title: dtColum,
        type: 'autocomplete',
        source: listaNome,
        width: '130px'
    });
}

var cellName1, cellName1, columRevert;
 
table = jspreadsheet(document.getElementById('spreadsheet'), {
    data:data,
    columns: colunas,
    toolbar:[
        {
            type: 'i',
            content: 'undo',
            onclick: function() {
                table.undo();
            }
        },
        {
            type: 'i',
            content: 'redo',
            onclick: function() {
                table.redo();
            }
        },
        {
            type: 'i',
            content: 'save',
            onclick: function () {
                table.download();
            }
        },
        {
            type: 'color',
            content: 'format_color_text',
            k: 'color'
        },
        {
            type: 'color',
            content: 'format_color_fill',
            k: 'background-color'
        },
        {
            type: 'i',
            content: 'data_object',
            onclick: function() {
                document.getElementById('console').innerHTML = JSON.stringify(data);
            }
        },
        /*{
            type: 'i',
            content: 'compress',
            onclick: function() {
                Merge();
            }
        },
        {
            type: 'i',
            content: 'replay_circle_filled',
            onclick: function() {
                RevertMerge();
            }
        }*/
    ],
    onselection: function(instance, x1, y1, x2, y2, origin) {
        cellName1 = jexcel.getColumnNameFromId([x1, y1]);
        cellName2 = jexcel.getColumnNameFromId([x2, y2]);
        if(x1 != undefined && x2 != undefined){
            columRevert = cellName1;
            console.log(cellName1)
        }
    },
    onchange: function(el, w, x, y, value, record){
        if(value != ''){
            cellOnChange = jexcel.getColumnNameFromId([x, y]);
            columRevert = alfabeto[x]+''+(y+1);
            DefinirEstiloPorNome(value);
        }
    }
});

function Merge(){
    var valorCell = document.getElementById('spreadsheet').jexcel.getLabel([cellName1]);

    if(cellName1 != undefined && cellName2 != undefined &&
       cellName1.substr(0, 1) == "A" && cellName2.substr(0, 1) == "A"){
        columRevert = cellName1;

        v1 = ConverterPraNumerico(cellName1);
        v2 = ConverterPraNumerico(cellName2);
        totalMerge = (v2-v1)+1;
        var cellPreenchida = 0, textValue = '';
        for (var i = 1; i <= totalMerge; i++) {
            var valorCell = table.getLabel(["A"+v1]);
            v1 ++;
            if(valorCell != '' && valorCell != textValue){
                cellPreenchida += 1;
                textValue = valorCell;
            }
            if(cellPreenchida > 1)
                return;
        }

        table.setMerge(cellName1, 1, totalMerge);
        table.setValue(cellName1, textValue, true);
    }
}

function ConverterPraNumerico(str){
    return str.replace(/[^0-9]/g, '');
}

function RevertMerge(){
    table.removeMerge(columRevert);
}

function setData(totalDias){
    var dataInicio = new Date();
    var dataFim = new Date();
    dataFim.setDate(dataFim.getDate() + totalDias);

    var difference= Math.abs(dataFim-dataInicio);
    return difference/(1000 * 3600 * 24);
}

function colorirPorNome(nome, tipo, cor){
    var numeroLinhas = JSON.stringify(data.length);
    for(var alfa = 1; alfa <= dias+2; alfa++) {
        coluna = alfabeto[alfa];
        for (var i = 1; i <= numeroLinhas; i++) {
            try {
                if(table.getLabel([coluna+i]) == nome && table.getLabel([coluna+i]) != '')
                    table.setStyle(coluna+i, tipo, cor);
            }
            catch (e) {
                break;
            }
        }
    }
}

function DefinirEstiloPorNome(nome){
    var verificarItem = ConjuntoEstilos.find( x => x.nome === nome);
    if(verificarItem == undefined && nome != ''){
        table.setStyle(columRevert, 'background-color', '#fff');
        table.setStyle(columRevert, 'color', '#000');

        var novoEstilo = new Estilo(nome, "#fff", "#000");
        ConjuntoEstilos.push(novoEstilo);
    }
    else if(verificarItem == undefined ||
            (verificarItem.background == '#fff' && 
            verificarItem.textColor == '#000')){
        table.setStyle(columRevert, 'background-color', '#fff');
        table.setStyle(columRevert, 'color', '#000');
    }
    else{
        if(verificarItem != undefined && verificarItem.background != '')
            table.setStyle(columRevert, 'background-color', verificarItem.background);
        if(verificarItem != undefined && verificarItem.textColor != '')
            table.setStyle(columRevert, 'color', verificarItem.textColor);
    }
}

function setEstilo(nome, k, v) {
    var buscarNomeEstilo = ConjuntoEstilos.find( x => x.nome === nome);

    switch (k) {
    case 'background-color':
        buscarNomeEstilo.background = v;
        break;
    case 'color':
        buscarNomeEstilo.textColor = v;
        break;
    default:
        break;
    }
}

function Estilo(nome, background, textColor) {
    this.nome = nome;
    this.background = background;
    this.textColor = textColor;
}