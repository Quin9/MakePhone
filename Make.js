// JavaScript implementation to generate required data
import fs from 'fs';
import path from 'path';
import xlsx from 'xlsx';
import { faker } from '@faker-js/faker';

// 解析当前文件夹路径
const __dirname = path.resolve(); // 使用 path.resolve() 确保路径正确

// 从命令行参数中获取输入文件路径和国家代码
const inputFilePath = process.argv[2];
const countryCode = process.argv[3];
if (!inputFilePath || !countryCode) {
    console.error('错误：请提供一个Excel文件作为参数以及国家代码（例如：DE, IC）。');
    process.exit(1);
}

// 检查文件是否存在
if (!fs.existsSync(inputFilePath)) {
    console.error(`错误：文件 "${inputFilePath}" 不存在。`);
    process.exit(1);
}

// 加载输入的Excel文件
const workbook = xlsx.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0]; // 默认读取第一个工作表
const inputData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

// 国家相关信息生成函数
function generateAddressAndPhone(countryCode) {
    const countryData = {
        BE: { streets: ["Rue Neuve", "Avenue Louise", "Chaussée de Charleroi", "Place Jourdan", "Boulevard Anspach", "Rue Royale", "Avenue de Tervuren", "Rue Belliard", "Boulevard du Midi", "Rue des Fripiers"], phoneCode: "+32" },
        BG: { streets: ["Tsarigradsko Shose", "Vitosha Boulevard", "Maria Luiza", "Dondukov Boulevard", "Rakovska", "Hristo Botev", "Cherni Vrah", "Alabin", "Graf Ignatiev", "Bulgaria Boulevard"], phoneCode: "+359" },
        CZ: { streets: ["Wenceslas Square", "Charles Square", "Na Prikope", "Nerudova Street", "Celetna Street", "Zizkov", "Holesovice", "Karlin", "Vinohrady", "Smichov"], phoneCode: "+420" },
        DK: { streets: ["Stroget", "Norrebrogade", "Gothersgade", "Vestergade", "Amagerbrogade", "Vesterbrogade", "Frederiksberg Alle", "Holmens Kanal", "Nyhavn", "Ostergade"], phoneCode: "+45" },
        DE: { streets: ["Alexanderplatz", "Berliner Str.", "Kurfuerstendamm", "Potsdamer Platz", "Unter den Linden", "Goethestr.", "Friedrichstr.", "Frankfurter Allee", "Karl-Marx-Str.", "Wilhelmstr."], phoneCode: "+49" },
        EE: { streets: ["Narva mnt", "Pärnu mnt", "Liivalaia", "Tartu mnt", "Rävala puiestee", "Estonia pst", "Ahtri", "Endla", "Järvevana tee", "Sõpruse pst"], phoneCode: "+372" },
        IE: { streets: ["O'Connell Street", "Grafton Street", "Dame Street", "Parnell Street", "Stephen's Green", "Pearse Street", "Capel Street", "George's Street", "Thomas Street", "Camden Street"], phoneCode: "+353" },
        EL: { streets: ["Ermou Street", "Patission Street", "Syngrou Avenue", "Academias Street", "Stadiou Street", "Panepistimiou Street", "Vas. Sofias Avenue", "Kifisias Avenue", "Vouliagmenis Avenue", "Peiraios Street"], phoneCode: "+30" },
        ES: { streets: ["Gran Via", "Calle de Alcalá", "Calle Serrano", "Calle de Atocha", "Paseo de la Castellana", "Calle de Velázquez", "Calle de Goya", "Calle Mayor", "Calle de Fuencarral", "Calle de Princesa"], phoneCode: "+34" },
        FR: { streets: ["Rue de Rivoli", "Avenue des Champs-Élysées", "Boulevard Haussmann", "Rue Saint-Honoré", "Rue de la Paix", "Avenue Montaigne", "Boulevard Saint-Germain", "Rue Cler", "Rue de Rennes", "Rue Mouffetard"], phoneCode: "+33" },
        HR: { streets: ["Ilica", "Zagrebačka avenija", "Savska cesta", "Heinzelova", "Maksimirska", "Palmotićeva", "Vukovarska", "Trpimirova", "Nova cesta", "Radnička cesta"], phoneCode: "+385" },
        IT: { streets: ["Via Roma", "Via Nazionale", "Corso Vittorio Emanuele", "Via del Corso", "Piazza Venezia", "Via Veneto", "Via Appia", "Via della Conciliazione", "Via XX Settembre", "Via Garibaldi"], phoneCode: "+39" },
        CY: { streets: ["Makarios Avenue", "Archbishop Street", "Stasikratous Street", "Ledra Street", "Agiou Andreou", "Larnaca Road", "Nikou Georgiou", "Tseriou Street", "Paphos Avenue", "Athinon Street"], phoneCode: "+357" },
        LV: { streets: ["Brīvības iela", "Elizabetes iela", "Valdemāra iela", "Krišjāņa Barona iela", "Dzirnavu iela", "Avotu iela", "Tallinas iela", "Miera iela", "Čaka iela", "Baznīcas iela"], phoneCode: "+371" },
        LT: { streets: ["Gedimino pr.", "Pilies g.", "Laisvės pr.", "Konstitucijos pr.", "Kalvarijų g.", "Ukmergės g.", "Šeimyniškių g.", "Žirmūnų g.", "Jasinskio g.", "Pylimo g."], phoneCode: "+370" },
        LU: { streets: ["Avenue de la Liberté", "Rue de Hollerich", "Rue de Strasbourg", "Boulevard Royal", "Rue des Capucins", "Rue Philippe II", "Avenue de la Gare", "Rue de Neudorf", "Rue des Bains", "Boulevard d'Avranches"], phoneCode: "+352" },
        HU: { streets: ["Andrássy út", "Rákóczi út", "Váci utca", "Kossuth Lajos utca", "Bartók Béla út", "Üllői út", "Nagykörút", "Bajcsy-Zsilinszky út", "Hegyalja út", "Alkotás út"], phoneCode: "+36" },
        MT: { streets: ["Republic Street", "Merchant Street", "St Paul Street", "St John Street", "Archbishop Street", "St George's Road", "Spinola Road", "Tower Road", "Triq Manwel Dimech", "Triq Il-Kbira"], phoneCode: "+356" },
        AT: { streets: ["Mariahilfer Str.", "Kärntner Str.", "Graben", "Landstraßer Hauptstr.", "Wiedner Hauptstr.", "Favoritenstr.", "Hütteldorfer Str.", "Gumpendorfer Str.", "Währinger Str.", "Alser Str."], phoneCode: "+43" },
        PL: { streets: ["Nowy Świat", "Marszałkowska", "Aleje Jerozolimskie", "Krakowskie Przedmieście", "Puławska", "Grzybowska", "Żwirki i Wigury", "Świętokrzyska", "Chmielna", "Złota"], phoneCode: "+48" },
        PT: { streets: ["Avenida da Liberdade", "Rua Augusta", "Rua Garrett", "Avenida Almirante Reis", "Rua da Prata", "Rua do Ouro", "Rua de São Bento", "Rua do Alecrim", "Rua dos Fanqueiros", "Rua da Palma"], phoneCode: "+351" },
        RO: { streets: ["Calea Victoriei", "Bulevardul Magheru", "Bulevardul Unirii", "Strada Lipscani", "Bulevardul Dacia", "Calea Dorobanți", "Strada Știrbei Vodă", "Calea Moșilor", "Strada Academiei", "Strada Franceză"], phoneCode: "+40" },
        SI: { streets: ["Slovenska cesta", "Tržaška cesta", "Dunajska cesta", "Celovška cesta", "Koprska ulica", "Šmartinska cesta", "Rožna dolina", "Vič", "Poljanska cesta", "Tivolska cesta"], phoneCode: "+386" },
        SK: { streets: ["Štefánikova", "Obchodná", "Záhradnícka", "Karadžičova", "Račianska", "Trnavská cesta", "Tomášikova", "Bajkalská", "Prievozská", "Einsteinova"], phoneCode: "+421" },
        FI: { streets: ["Mannerheimintie", "Aleksanterinkatu", "Esplanadi", "Kaivokatu", "Lönnrotinkatu", "Runeberginkatu", "Fredrikinkatu", "Lapinlahdenkatu", "Kalevankatu", "Annankatu"], phoneCode: "+358" },
        SE: { streets: ["Drottninggatan", "Sveavägen", "Kungsgatan", "Hamngatan", "Vasagatan", "Stureplan", "Hornsgatan", "Götgatan", "Regeringsgatan", "Skeppsbron"], phoneCode: "+46" },
        UK: { streets: ["Oxford Street", "Regent Street", "Bond Street", "Piccadilly", "The Strand", "King's Road", "Fleet Street", "Baker Street", "Whitehall", "High Holborn"], phoneCode: "+44" },
        NL: { streets: ["Damstraat", "Kalverstraat", "Leidsestraat", "Rokin", "Prinsengracht", "Haarlemmerdijk", "Nieuwezijds Voorburgwal", "Vijzelstraat", "Lijnbaansgracht", "Kinkerstraat"], phoneCode: "+31", },
    };

    const data = countryData[countryCode.toUpperCase()];
    if (!data) {
        console.error('错误：不支持的国家代码。');
        process.exit(1);
    }

    return {
        generateAddress: () => {
            const street = data.streets[Math.floor(Math.random() * data.streets.length)];
            const number = Math.floor(Math.random() * 200) + 1;
            return `${street} ${number}`;
        },
        generatePhone: () => `${data.phoneCode} ${Math.floor(1000000 + Math.random() * 9000000)}`
    };
}

const { generateAddress, generatePhone } = generateAddressAndPhone(countryCode);

// 生成邮箱
function generateEmail(lastName, firstName) {
    const randomNumbers = Math.floor(100 + Math.random() * 900);
    return `${lastName}${firstName}${randomNumbers}@mail.com`.replace(/[^a-zA-Z0-9@.]/g, '');
}

// 处理输入数据
const outputData = inputData.map((row, index) => {
    if (row.length < 4) {
        console.error(`第 ${index + 1} 行数据不完整，跳过处理。`);
        return null;
    }

    const lastName = row[0]; // 第一列
    const firstName = row[1]; // 第二列
    const birthday = row[2] ? new Date(row[2]) : new Date(); // 第三列
    const bankInfo = row[3]; // 第四列

    const email = generateEmail(lastName, firstName);
    const formattedBirthday = birthday.toLocaleDateString('nl-NL', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
    });
    const phone = generatePhone();
    const address = generateAddress();
    const password = "Aa123123@";

    return {
        姓: lastName,
        名: firstName,
        电话: phone,
        邮箱: email,
        密码: password,
        生日: formattedBirthday,
        街道地址: address,
        银行信息: bankInfo,
    };
}).filter(Boolean);

// 生成随机文件名
const randomNumbers = Math.floor(100 + Math.random() * 900);
const outputFileName = `${countryCode}${randomNumbers}.xlsx`;
const outputFilePath = path.join(__dirname, outputFileName);
console.log(`尝试保存文件到: ${outputFilePath}`);

// 保存输出数据到新Excel文件
try {
    const outputSheet = xlsx.utils.json_to_sheet(outputData);
    const outputWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(outputWorkbook, outputSheet, 'Output');
    xlsx.writeFile(outputWorkbook, outputFileName); // 直接写入文件名，无需带完整路径
    console.log(`数据处理完成。文件已保存到: ${outputFileName}`);
} catch (error) {
    console.error('保存文件时出错:', error.message);
    process.exit(1);
}
