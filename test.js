const xlsx = require("xlsx");
let data = [
    {
        name: "Bhuvanesh",
        age: 10,
        subs: ["p","c","m"],
        sports: [
            {
                name: "Chess",
                overall_score: 7.8,
                games: [
                    {
                        name: "Jalgaon Open",
                        price: 100                           
                    },
                    {
                        name: "Jalgaon Open",
                        price: 100
                    }
                ]
            },
            {
                name: "Cricket",
                overall_score: 5,
                games: [
                    {
                        name: "Jalgon open",
                        price: 100
                    }
                ]
            }
        ]
    },
    {
        name: "Mitesh",
        age: 11,
        subs: ["p","c","m","b"],
        sports: [
            {
                name: "Football",
                overall_score: 7,
                games: [
                    {
                        name: "Jalgon open",
                        price: 100
                    }
                ]
            } 
        ]
    },
]
// const data = [{
//     name: "Bhuvanesh",
//     age: "10",
//     address: {
//         street: "Shivdarshan",
//         district: "Pune",
//         state: "Maharashtra"
//     },
//     books: [
//         {
//             name: "Charles 1",
//             ISBN: "11123-23422"
//         },
//         {
//             name: "Charles 2",
//             ISBN: "11123-10234"
//         }
//     ],
//     games: [
//         {
//             category: "Cat 1",
//             games: [
//                 {
//                     name: "GTA 5",
//                     price: 5000
//                 },
//                 {
//                     name: "RDR 2",
//                     price: 1000
//                 }
//             ]
//         },
//         {
//             category: "Cat 2",
//             games: [
//                 {
//                     name: "CSGO",
//                     price: 0
//                 }
//             ]
//         }
//     ]
// }
// ]

const fs = require("fs");
// let data = null;
async function readAndParseJson(path) {
    try {
      const data = await fs.promises.readFile(path, 'utf-8');
      return JSON.parse(data);
    } catch (error) {
      console.error(error);
    }
  }
  
(async () => {
// data = await readAndParseJson('input.json');
const getLetter = (num) => {
    return String.fromCharCode(num + 64);
}

const process_object = (object, default_sheet_name, object_index) => {
    let worksheets = [];
    console.log(`Entry -> ${default_sheet_name}`)
    if(Array.isArray(object)){
        let json_list = [];
        let array_worksheets = [];
        console.log(`Type -> array`)
        for(let i of object){
            const sheets = process_object(i,`${default_sheet_name}.${object.indexOf(i)}`,object.indexOf(i)+2);
            array_worksheets = array_worksheets.concat(sheets);
        }
        
        json_list = array_worksheets.map(list => xlsx.utils.sheet_to_json(list[0]));
        console.log(`JSON Dump for ${default_sheet_name}`);
        console.log(json_list);

        // json_list
        let can_concatenate = [];
        for(let i=0;i<json_list.length;i++){
            if(json_list[i].length == 1){
                can_concatenate.push([json_list[i][0],i]);
            }else{
                worksheets.push(array_worksheets[i]);
            }
        }

        console.log(`Concatenations for ${default_sheet_name}`);
        console.log(can_concatenate); 
        console.log(`Worksheets at ${default_sheet_name}`);
        console.log(worksheets);
        
        let con = can_concatenate.reduce((acc,obj,ind) => {
            let key = JSON.stringify(Object.keys(obj[0]));
            if(!acc[key]){
                acc[key] = [[obj[0]],array_worksheets[obj[1]][2],1,obj[1]];
            }else{
                acc[key][0].push(obj[0]);
                acc[key][1].push(array_worksheets[obj[1]][2]);
                acc[key][2] = acc[key][2]+1;
                acc[key][3] = obj[1];
            }
            return acc;
        },{});

        console.log(`Commons for ${default_sheet_name}`);
        console.log(con);

        Object.keys(con).forEach(key => {
            if(con[key][2] > 1){
                const sheet = xlsx.utils.json_to_sheet(con[key][0]);
                worksheets.push([sheet, default_sheet_name, con[key][1]]);
            }else{
                worksheets.push(array_worksheets[con[key][3]]);
            }
        });

        console.log(`Worksheets at ${default_sheet_name}`);
        console.log(worksheets);

        return worksheets;
    }else{
        let hyperlinks = [];
        let keys = Object.keys(object)
        for(let i=0; i<keys.length; i++) {
            if(Array.isArray(object[keys[i]]) && typeof(object[keys[i]][0]) != "object"){
                object[keys[i]] = object[keys[i]].join(",");
            }else
            if(typeof(object[keys[i]]) == "object" && object[keys[i]] != null){
                const sheet_name =`${default_sheet_name}.${keys[i]}`;
                const sheets = process_object(object[keys[i]], sheet_name, object_index);
                worksheets = worksheets.concat(sheets);
                hyperlinks.push([i+1,object_index,sheet_name]);
                object[keys[i]] = sheet_name;
            }
        }
        const obj_sheet = xlsx.utils.json_to_sheet([object]);
        // if(hyperlinks.length){
        //     for(let i=0;i<hyperlinks.length;i++){
        //         console.log(`${getLetter(hyperlinks[i][0])}${hyperlinks[i][1]}`);
        //         obj_sheet[`${getLetter(hyperlinks[i][0])}${hyperlinks[i][1]}`].l = { Target: `#${hyperlinks[i][2]}!A1` };
        //     }
        // }
        worksheets.push([obj_sheet, default_sheet_name, hyperlinks]);
        console.log(`Worksheets for ${default_sheet_name}`);
        console.log(worksheets);
        return worksheets;
    }
}

// xlsx.utils.cell_set_hyperlink()

const wb = xlsx.utils.book_new();
// const ws = xlsx.utils.json_to_sheet(data);
const sheets = process_object(data,"Sheet1",1);
console.log("Final Sheets")
console.log(sheets);
sheets.forEach(sheet => {
    console.log("Hyperlinks----");
    console.log(sheet[2]);
    let hyperlinks = sheet[2];
    for(let j=0;j<hyperlinks.length;j++){
        if(Array.isArray(hyperlinks[j][0])){
            hyperlinks[j] = hyperlinks[j][0];
        }
    }
    hyperlinks = hyperlinks.filter((item) => item.length != 0);
    for(let j=0;j<hyperlinks.length;j++){
        console.log(`${getLetter(hyperlinks[j][0])}${hyperlinks[j][1]}`);
        // if(!sheet[0][`${getLetter(hyperlinks[j][0])}${hyperlinks[j][1]}`]) continue;
        sheet[0][`${getLetter(hyperlinks[j][0])}${hyperlinks[j][1]}`].l = { Target: `#${hyperlinks[j][2]}!A1` };
    }
})
sheets.forEach(sheet => xlsx.utils.book_append_sheet(wb,sheet[0],sheet[1]));
// xlsx.utils.book_append_sheet(wb,ws);
xlsx.writeFile(wb,"out.xlsx");
})();

