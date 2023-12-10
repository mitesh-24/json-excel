const xlsx = require("xlsx");
let data = require("./main.js");
// If object is found process_object
// If array is found proces_array
// process_object -> returns -> [[sheet, sheet_name], depending_sheets, hyperlinks]

// process_object takes (object, object_key, depth)
// process_array takes (object, object_key)

const getLetter = (num) => {
    return String.fromCharCode(num + 64);    
}

const process_object = (object, object_key, depth) => {
    console.log("In process object");
    console.log(object);

    let sheet = null;
    let depending_sheets = [];
    let hyperlinks = [];

    // if(Array.isArray(object)){
    //    return;
    // }

    let keys = Object.keys(object);
    
    for(let i=0;i<keys.length;i++){
        const obj_sheet_name = `${object_key}.${keys[i]}`;
        if(Array.isArray(object[keys[i]])){
            if((object[keys[i]].length && typeof(object[keys[i]][0]) != "object")){
                object[keys[i]] = object[keys[i]].join(",");
                continue;
            }
            let [obj_sheet, obj_depending_sheet] = process_array(object[keys[i]], obj_sheet_name);
            depending_sheets.push(obj_sheet);
            depending_sheets = depending_sheets.concat(obj_depending_sheet);
            hyperlinks.push([obj_sheet_name, `${getLetter(i+1)}${depth}`]);

            object[keys[i]] = obj_sheet_name;
        }else
        if(typeof(object[keys[i]]) == "object" && object[keys[i]] != null){            
            let [obj_sheet, obj_depending_sheet, obj_hyperlinks] = process_object(object[keys[i]], obj_sheet_name, depth);
            depending_sheets = depending_sheets.concat(obj_depending_sheet);
            hyperlinks = hyperlinks.concat(obj_hyperlinks);

            depending_sheets.push(obj_sheet);
            hyperlinks.push([obj_sheet_name, `${getLetter(i+1)}${depth}`]);

            object[keys[i]] = obj_sheet_name;
            console.log(`Hyperlinks generated for object ${obj_sheet_name}`);
            console.log(hyperlinks);
        }
    }
    
    console.log(`Object ${object_key} after processing`);
    console.log(object);

    sheet = xlsx.utils.json_to_sheet([object]);

    for(let i of hyperlinks){
        if(!sheet[i[1]]) continue;
        sheet[i[1]].l = {Target: `#${i[0]}!A1`} 
    }

    return [[sheet,object_key], depending_sheets, hyperlinks];

}

const process_array = (object, object_key) => {
    console.log("In process array -> ");
    console.log(object);

    let sheet = null;
    let depending_sheets = [];
    let obj_sheets = [];
    let hyperlinks = [];

    for(let i=0;i<object.length;i++){
        let sheet_key = `${object_key}.${i}`;
        let [obj_sheet, obj_depending_sheet, obj_hyperlinks] = process_object(object[i], sheet_key, i+2);
        obj_sheets.push(obj_sheet);
        depending_sheets = depending_sheets.concat(obj_depending_sheet);
        hyperlinks = hyperlinks.concat(obj_hyperlinks);
    }
    console.log(`Sheets generated for ${object_key}`);
    console.log(obj_sheets);

    obj_sheets = obj_sheets.map(sheet => xlsx.utils.sheet_to_json(sheet[0])[0]);
    console.log(`JSON dump for ${object_key}`);
    console.log(obj_sheets);

    sheet = xlsx.utils.json_to_sheet(obj_sheets);

    console.log(`Hyperlinks for ${object_key}`);
    console.log(hyperlinks);

    for(let i of hyperlinks){
        sheet[i[1]].l = {Target: `#${i[0]}!A1`}
    }

    console.log(`Sheet for ${object_key}`);
    console.log(sheet);

    return [[sheet, object_key], depending_sheets];
}
if(!Array.isArray(data)) data = [data];

let [sheet, depending_sheets] = process_array(data, "Sheet1");

console.log("Final Sheets");
console.log(sheet);
console.log(depending_sheets);

const wb = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(wb, sheet[0], sheet[1]);
for(let d of depending_sheets){
    xlsx.utils.book_append_sheet(wb, d[0], d[1]);
}

xlsx.writeFile(wb,"out_new.xlsx");
