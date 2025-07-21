const express = require('express');
// import { Request, Response } from "express";
const xlsx = require('xlsx');
// const yahooFinance = require('yahoo-finance2');
import yahooFinance from 'yahoo-finance2';
const fs = require('fs');
yahooFinance.suppressNotices(['yahooSurvey']);

type stockInfo = {
    No: number,
    Particulars: string,
    Quantity: number,
    Investment: number,
    code: string,
    portfolioPercentage: number,
    cmp: number,
    presentValue: number,
    gainLoss: number,
    peRatio: number,
    latestEarnings: number,
    industry: string
}



// const app = express();
// const port = 8080;
// express.json();
let stocks: stockInfo[] = [];
let stockDetails: any[] = [];

const getValidSymbol = async (name: string) => {
    const results = await yahooFinance.search(name);
    if (!results.quotes || results.quotes.length === 0) return "";

    // Find the first quote object that has a 'symbol' property
    const match = results.quotes.find(quote => 'symbol' in quote)?.symbol;

    return match ?? "";
}

const getStockData = async (stock: stockInfo) => {

    let searchName: string = stock.code.toString();
    if (/^[A-Za-z]+$/.test(searchName)) {
        searchName = searchName + '.NS';
    }
    else searchName = await getValidSymbol(stock.Particulars);

    try {
        const quote = await yahooFinance.quote(searchName);
        const profile = await yahooFinance.quoteSummary(searchName, { modules: ['assetProfile'] });

        stock.cmp = quote.regularMarketPrice ?? -1;
        stock.peRatio = quote.trailingPE ?? -1;
        stock.latestEarnings = quote.epsTrailingTwelveMonths ?? -1;
        stock.industry = profile.assetProfile?.industry ?? 'N/A';
    }
    catch (error: any) {
        // console.log(error.message);
    }

    return stock;
}

const saveStockDetails = () => {

    fs.writeFile('stockDetails.json', JSON.stringify(stockDetails, null, 2), (err: any) => {
        if (err) {
            console.error('Error writing file:', err);
        } else {
            console.log('File saved successfully.');
        }
    });

}

(async () => {
    try {
        const data: any = xlsx.readFile('data.xlsx');
        const sheetName = data.SheetNames[0];
        const worksheet = data.Sheets[sheetName];

        let rawData: any[] = xlsx.utils.sheet_to_json(worksheet, { defval: "", range: 1 });

        for (const data of rawData) {
            if (data.No === '') continue;

            let newStock: stockInfo = {
                No: data.No,
                Particulars: data.Particulars,
                Quantity: data.Qty,
                Investment: data.Investment,
                code: data['NSE/BSE'],
                portfolioPercentage: data['Portfolio(%)'],
                cmp: -1,
                presentValue: -1,
                gainLoss: -1,
                peRatio: -1,
                latestEarnings: -1,
                industry: 'N/A',
            }

            stocks.push(newStock);
        }
        console.log("size of stocks is ", stocks.length);

        stockDetails = await Promise.all(stocks.map(async (stock: stockInfo) => {
            try {
                return await getStockData(stock);
            } catch (e) {
                console.error(`Failed for stock ${stock.code}`);
                return stock;
            }
        }
        ))

        // console.log(stockDetails);

        saveStockDetails();
    }
    catch (error: any) {
        //console.log(error);
    }
})()



// app.get('/', (req: Request, res: Response) => {
//     res.send('Hello World!');
// });


// app.listen(port, () => {
//     //console.log(`Example app listening on port ${port}`);
// })