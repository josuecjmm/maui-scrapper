const axios = require('axios')
const cheerio = require('cheerio')
const fs = require('fs');
const json2xls = require('json2xls');
const xlsx = require('./xlsx');

async function loadPage(url) {
    try {
        const response = await axios.get(url);
        const html = response.data;
        const $ = cheerio.load(html);

        const name = $('.p-campaign-header h1').text();
        const totalRaised = $('.o-campaign-sidebar-wrapper .hrt-disp-inline').text();
        let goal = $('.o-campaign-sidebar-wrapper .hrt-text-body-sm').text()
        if (goal === 'USD raised') {
            goal = 'N/A'
        } else {
            goal = goal.split('of ')[1].split(' ')[0];
        }
        let totalDonations = $('.hrt-text-gray-dark').text()
        const description = $('.o-campaign-story').text()
        return {
            'Clean Link': url,
            Name: name,
            Raised: parseInt(totalRaised.replaceAll('$', '').replaceAll(',', '')),
            'Goal': parseInt(goal.replaceAll('$', '').replaceAll(',', '')),
            Donors: parseInt(totalDonations),
            'Goal %': '',
            'Goal Met': '',
            'Top Donation': '',
            Notes: '',
            Description: description.replaceAll(/\n/g, '').slice(0, 150) + '...'
        }
    } catch (e) {
        console.error(e);
        return {
            'Clean Link': url,
            Name: '',
            Raised: '',
            'Goal': '',
            Donors: '',
            'Goal %': '',
            'Goal Met': '',
            'Top Donation': '',
            Notes: '',
            Description: ''
        }
    }
}

async function extractDataFromOriginalSheet() {
    const originalFileName = './data/GoFundMe_Data.xlsx'
    const file = xlsx.getFileContent(originalFileName)
    const sheet = xlsx.getSheetContent(file, 'CleanData')
    const columnValues = xlsx.getColumnValues(sheet, 'C');

    let formattedLinks = columnValues.map(link => {
        return link.replaceAll(/<t>/g, '').replaceAll('</t>', '')
    })
    formattedLinks.shift();
    return formattedLinks
}

async function multipleDonations() {
    const links = await extractDataFromOriginalSheet();
    const data = [];
    let i = 1;
    for await (const url of links) {
        console.log('#', i, ' Extracting data from=> ', url)
        const extractedData = await loadPage(url);
        data.push(extractedData)
        i++;
    }
    return data;
}

async function createExcelFile() {
    const data = await multipleDonations();
    const filename = 'extracted_donations_data.xlsx';
    const xls = json2xls(data);
    fs.writeFileSync(filename, xls, 'binary', (err) => {
        if (err) {
            console.log("writeFileSync :", err);
        }
        console.log(filename + " file is saved!");
    });
}

createExcelFile().then().catch()
