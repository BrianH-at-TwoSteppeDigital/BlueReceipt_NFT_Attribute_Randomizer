const reader = require('xlsx');
let json = [];

async function randomCode() {
    try {
        let character = ["Male", "Female"]
        let hair_color = ["Brown", "Blonde", "Red", "Black", "Grey"];
        let height = ["6'3", "6'1", "5 '9", "5'5"];
        let skin = ["Yes", "No"];
        let skin_color = ["Red", "Blue", "Camo", "Gold", "Platinum"];
        let accessory = ["Hat", "Earning", "Necklace", "Glasses"];
        let weapon = ["Weapon", "No weapon"];
        let weapon_gun = ["Hand gun", "semi automatic", "automatic", "rocket launcher", "sniper"];
        let weapon_bonus = ["1 Rounds", "5 Rounds", "20 Rounds", "100 Rounds"];
        let vehicle = ["Vehicle", "no vehicle"];
        let vehicle_car = ["BMX", "Scooter", "Junk Car", "Super Car", "Tank"];
        let vehicle_bonus = ["Rims", "Spoiler", "straight pipe", "nitro"];
        let superpower = ["Yes", "No"];
        let superpower_bonus = ["Invisibility", "super speed", "super jump", "Max health","Flying"];
        let superpower_percent = ["1%", "10%", "50%", "100%"];

        for (var i = 0; i < 10000; i++) {
            let codeStr = "";

            let x1 = Math.floor((Math.random() * 2));
            let x2 = Math.floor((Math.random() * 5));
            let x3 = Math.floor((Math.random() * 4));
            let x4 = Math.floor((Math.random() * 2));
            let x5 = Math.floor((Math.random() * 5));
            let x6 = Math.floor((Math.random() * 4));
            let x7 = Math.floor((Math.random() * 2));
            let x8 = Math.floor((Math.random() * 5));
            let x9 = Math.floor((Math.random() * 4));
            let x10 = Math.floor((Math.random() * 2));
            let x11 = Math.floor((Math.random() * 5));
            let x12 = Math.floor((Math.random() * 4));
            let x13 = Math.floor((Math.random() * 2));
            let x14 = Math.floor((Math.random() * 5));
            let x15 = Math.floor((Math.random() * 4));

            await json.push({});

            json[i].CODE = codeStr;
            json[i].CHARACTER = character[x1];
            json[i].HAIRCOLOR = hair_color[x2];
            json[i].HEIGHT = height[x3];
            if(json[i].CHARACTER == "Male") {
                codeStr = "Up;" + (x2+1) + ";" + direction(x3);
            } else if(json[i].CHARACTER == "Female") {
                codeStr = "Down;" + (x2+1) + ";" + direction(x3);
            }
            json[i].SKIN = skin[x4];
            if(json[i].SKIN == "Yes") {
                json[i].SKIN_COLOR = skin_color[x5];
                json[i].ACCESSORY = accessory[x6];
                codeStr = codeStr + "Up;" + (x5+1) + ";" + direction(x6);
            } else if(json[i].SKIN == "No") {
                json[i].SKIN_COLOR = "";
                json[i].ACCESSORY = "";
                codeStr = codeStr + "Down;";
            }
            json[i].WEAPON = weapon[x7];
            if(json[i].WEAPON == "Weapon") {
                json[i].WEAPON_GUN = weapon_gun[x8];
                json[i].WEAPON_BONUS = weapon_bonus[x9];
                codeStr = codeStr + "Up;" + (x8+1) + ";" + direction(x9);
            } else if(json[i].WEAPON == "No weapon") {
                json[i].WEAPON_GUN = "";
                json[i].WEAPON_BONUS = "";
                codeStr = codeStr + "Down;";
            }
            json[i].VEHICLE = vehicle[x10];
            if(json[i].VEHICLE == "Vehicle") {
                json[i].VEHICLE_CAR = vehicle_car[x11];
                json[i].VEHICLE_BONUS = vehicle_bonus[x12];
                codeStr = codeStr + "Up;" + (x11+1) + ";" + direction(x12);
            } else if(json[i].VEHICLE == "no vehicle") {
                json[i].VEHICLE_CAR = "";
                json[i].VEHICLE_BONUS = "";
                codeStr = codeStr + "Down;";
            }
            json[i].SUPERPOWER = superpower[x13];
            if (json[i].SUPERPOWER == "Yes") {
                json[i].SUPERPOWERBONUS = superpower_bonus[x14];
                json[i].PERCENT = superpower_percent[x15];
                codeStr = codeStr + "Up;" + (x14+1) + ";" + direction(x15);
            } else if(json[i].SUPERPOWER == "No") {
                json[i].SUPERPOWERBONUS = "";
                codeStr = codeStr + "Down;";
            }
            json[i].CODE = codeStr;

            if (i == 9999) {
                console.log("---- Excel ---  Save --- Finish ----");
            }
        }

        let workBook = await reader.utils.book_new();
        const workSheet = await reader.utils.json_to_sheet(json);
        await reader.utils.book_append_sheet(workBook, workSheet, `response`);
        let exportFileName = `./excel/response.xls`;
        await reader.writeFile(workBook, exportFileName);

        for (var i = 0; i < 10000; i++) {
            let codeStr = "";

            let x1 = Math.floor((Math.random() * 2));
            let x2 = Math.floor((Math.random() * 5));
            let x3 = Math.floor((Math.random() * 4));
            let x4 = Math.floor((Math.random() * 2));
            let x5 = Math.floor((Math.random() * 5));
            let x6 = Math.floor((Math.random() * 4));
            let x7 = Math.floor((Math.random() * 2));
            let x8 = Math.floor((Math.random() * 5));
            let x9 = Math.floor((Math.random() * 4));
            let x10 = Math.floor((Math.random() * 2));
            let x11 = Math.floor((Math.random() * 5));
            let x12 = Math.floor((Math.random() * 4));
            let x13 = Math.floor((Math.random() * 2));
            let x14 = Math.floor((Math.random() * 5));
            let x15 = Math.floor((Math.random() * 4));

            await json.push({});

            json[i].CODE = codeStr;
            json[i].CHARACTER = character[x1];
            json[i].HAIRCOLOR = hair_color[x2];
            json[i].HEIGHT = height[x3];
            if(json[i].CHARACTER == "Male") {
                codeStr = "Up;" + (x2+1) + ";" + direction(x3);
            } else if(json[i].CHARACTER == "Female") {
                codeStr = "Down;" + (x2+1) + ";" + direction(x3);
            }
            json[i].SKIN = skin[x4];
            if(json[i].SKIN == "Yes") {
                json[i].SKIN_COLOR = skin_color[x5];
                json[i].ACCESSORY = accessory[x6];
                codeStr = codeStr + "Up;" + (x5+1) + ";" + direction(x6);
            } else if(json[i].SKIN == "No") {
                json[i].SKIN_COLOR = "";
                json[i].ACCESSORY = "";
                codeStr = codeStr + "Down;";
            }
            json[i].WEAPON = weapon[x7];
            if(json[i].WEAPON == "Weapon"){
                json[i].WEAPON_GUN = weapon_gun[x8];
                json[i].WEAPON_BONUS = weapon_bonus[x9];
                codeStr = codeStr + "Up;" + (x8+1) + ";" + direction(x9);
            } else if(json[i].WEAPON == "No weapon") {
                json[i].WEAPON_GUN = "";
                json[i].WEAPON_BONUS = "";
                codeStr = codeStr + "Down;";
            }
            json[i].VEHICLE = vehicle[x10];
            if(json[i].VEHICLE == "Vehicle") {
                json[i].VEHICLE_CAR = vehicle_car[x11];
                json[i].VEHICLE_BONUS = vehicle_bonus[x12];
                codeStr = codeStr + "Up;" + (x11+1) + ";" + direction(x12);
            } else if(json[i].VEHICLE == "no vehicle") {
                json[i].VEHICLE_CAR = "";
                json[i].VEHICLE_BONUS = "";
                codeStr = codeStr + "Down;";
            }
            json[i].SUPERPOWER = superpower[x13];
            if (json[i].SUPERPOWER == "Yes") {
                json[i].SUPERPOWERBONUS = superpower_bonus[x14];
                json[i].PERCENT = superpower_percent[x15];
                codeStr = codeStr + "Up;" + (x14+1) + ";" + direction(x15);
            } else if(json[i].SUPERPOWER == "No") {
                json[i].SUPERPOWERBONUS = "";
                codeStr = codeStr + "Down;";
            }
            json[i].CODE = codeStr;

            if (i == 9999) {
                console.log("---- Excel_1 ---  Save --- Finish ----");
            }
        }

        let workBook1 = await reader.utils.book_new();
        const workSheet1 = await reader.utils.json_to_sheet(json);
        await reader.utils.book_append_sheet(workBook1, workSheet1, `response`);
        let exportFileName1 = `./excel/response1.xls`;
        await reader.writeFile(workBook1, exportFileName1);

        function direction(x) {
            switch (x) {
                case 0:
                    return "North;";
                case 1:
                    return "South;";
                case 2:
                    return "East;";
                case 3:
                    return "West;";
                default:
                    return "";
            }
        }

    } catch(error) {
        console.log(error);
    }
}

randomCode();