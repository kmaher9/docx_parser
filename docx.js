const StreamZip = require('node-stream-zip');

/**
 * Extracts raw XML data from a DOCX file
 * @param {string} filePath The location of the .docx file
 * @returns {Promise<string>}
 */
const open = (filePath) => new Promise((resolve, reject) => {

    // Instantiate zip stream
    const zip = new StreamZip({
        file: filePath,
        storeEntries: true
    })

    zip.on('ready', () => {

        // Uint8Array[]
        let chunks;

        // Buffer
        let content;

        // Open enclosed XML file
        zip.stream('word/document.xml', (err, stream) => {

            if (err) return reject(err);
            if (!stream) return reject('No stream.');

            // Append data chunks
            stream.on('data', chunk => chunks.push(chunk));

            // Handle stream close
            stream.on('end', () => {
                content = Buffer.concat(chunks);
                zip.close();
                resolve(content.toString());
            });

        })
    })

})

/**
 * Extracts text from an DOCX file
 * @param {string} filePath The location of the .docx file
 * @returns {Promise<string>}
 */
const extract = (filePath) => new Promise((resolve, reject) => {

    open(filePath)

        .then((result) => {

            // Body text
            let body = '';

            // Funky regex magic
            let components = result.split('<w:t');
            for (let i = 0; i < components.length; i++) {

                let tags = components[i].split('>');
                let content = tags[1].replace(/<.*$/, "");

                body += content + ' ';

            }

            resolve(body)

        })
        
        .catch(reject)

});

module.exports = {

    extract, open

}
