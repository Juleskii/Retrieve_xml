import XLSX from 'xlsx';
import xml2js from 'xml2js';
import fs from 'fs';
import axios from 'axios';

const workbook = XLSX.readFile('2019_Agency_Info.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(worksheet);

// Get links from worksheet and store in array
const getLinks = () => {
  const links = [];
  data.forEach((row) => {
    const keys = Object.keys(row);
    const lastCell = row[keys[keys.length - 1]];
    const link = lastCell.substring(lastCell.indexOf('http'));
    links.push(link);
  });

  // Returns 141 links
  return links;
};

// Iterate through links and fetch XML data
const fetchXML = async () => {
  const firstLink = getLinks()[0];
  const response = await axios.get(firstLink);
  const xml = response.data; // Assign the response data to 'xml'
  return xml;
};

// Parse XML data and return JSON
const parseXML = async () => {
  const xml = await fetchXML();
  const parser = new xml2js.Parser();
  const json = await parser.parseStringPromise(xml);
  return json;
};

parseXML().then((json) => {
  let count = 0;
  // Loop through the json.rss.channel[0].item array
  // and log each property
  json.rss.channel[0].item.forEach((item) => {
    const {
      title,
      link,
      pubDate,
      jobId,
      location,
      state,
      department,
      advertiseFromDate,
    } = item;

    console.log(item);

    // Log all variables
    console.log('title: ', title[0]);
    console.log('link: ', link[0]);
    console.log('pubDate: ', pubDate[0]);
    console.log('jobId: ', jobId);
    console.log('location: ', location);
    console.log('state: ', state);
    console.log('department: ', department);
    console.log('advertiseFromDate: ', advertiseFromDate);
    console.log('-----------------------------------');
  });
});
