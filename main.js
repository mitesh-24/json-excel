const fs = require("fs");
const data = [
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

// const rawData = fs.readFileSync("./input.json", 'utf8');
// const data = JSON.parse(rawData);

module.exports = data;
// sheet
// rows
// [{}]