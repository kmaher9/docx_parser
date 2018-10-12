var docx = require('./docx')

docx.extract('path\\to\\my\\file.docx').then(function(res, err) {
    if (err) {
        console.log(err)
    }
    console.log(res)
})