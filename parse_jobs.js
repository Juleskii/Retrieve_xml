(async () => {
  try {
    const XLSX = require('xlsx');
    const xml2js = require('xml2js');
    const fs = require('fs');
    const axios = require('axios');

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
      return links;
    };

    const links = getLinks();
    const results = []; // Array to store JSON objects

    // Function to fetch and parse XML for a single link
    const processLink = async (link) => {
      try {
        const json = await parseXML(link);

        // Loop through the json.rss.channel[0].item array
        // and create JSON objects
        json.rss.channel[0].item.forEach((item) => {
          const { title, link, pubDate } = item;
          const jobId = item['joblisting:jobId'][0];
          const location = item['joblisting:location'][0];
          const state = item['joblisting:state'][0];
          const department = item['joblisting:department'][0];
          const advertiseFromDate =
            item['joblisting:advertiseFromDate'][0];

          // const createdTime = new Date(pubDate[0]).toISOString();
          const inputDate = new Date(pubDate[0]);
          const createdTime = new Date(inputDate);
          const formattedDate = inputDate
            .toLocaleDateString('en-US', {
              year: 'numeric',
              month: '2-digit',
              day: '2-digit',
            })
            .replace(/\//g, '-');

          // Create JSON object
          const jsonObject = {
            title: title[0],
            jobUrl: link[0],
            createdTime,
            created: formattedDate,
            jobId,
            location,
            state,
            department,
            advertiseFromDate,
          };

          results.push(jsonObject); // Push the JSON object to the results array
        });
      } catch (error) {
        console.error('Error processing link:', link);
        console.error(error);
      }
    };

    // Function to fetch and parse XML for a single link
    const parseXML = async (link) => {
      const xml = await fetchXML(link);
      const parser = new xml2js.Parser();
      const json = await parser.parseStringPromise(xml);
      return json;
    };

    // Function to fetch XML for a single link
    const fetchXML = async (link) => {
      const response = await axios.get(link);
      const xml = response.data;
      return xml;
    };

    // Use Promise.all to process all links concurrently
    await Promise.all(links.map(processLink));

    // Log the results array containing JSON objects
    const jsonResults = JSON.stringify(results, null, 2);
    fs.writeFileSync('results.json', jsonResults);
    console.log('Results saved to results.json');
  } catch (error) {
    console.error('An error occurred:');
    console.error(error);
  }
})();
