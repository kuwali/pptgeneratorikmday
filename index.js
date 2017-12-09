const Bluebird = require('bluebird');
const pptx = require('pptxgenjs');
const fs = Bluebird.promisifyAll(require('fs'));
const csv = require('csvtojson');
const cheerio = require('cheerio');
const request = Bluebird.promisifyAll(require('superagent'));
require('superagent-retry-delay')(request);

pptx.setLayout('LAYOUT_WIDE');

var slide = pptx.addNewSlide();
var counter = 0;

const cetak = (rows) => {
  return Bluebird.resolve().then(() => {
    return request
      .get(`http://kawung.mhs.cs.ui.ac.id/~m.prakash/studentsearch/index.php?npm=${rows[2]}`)
      .retry(2, 3000, [401, 404])
      .then(response => {
        var $ = cheerio.load(response.text);
        console.log(counter++);
        return pptx.addNewSlide()
          .addImage({
            x: 5.2,
            y: 1,
            w: 3,
            h: 3,
            data: $('img').attr('src')
          })
          .addText(rows[0], {
            x: 0.2,
            y: 4.2,
            w: 13,
            h: 1,
            align: 'c',
            font_face: 'Montserrat',
            color: '76c9d9',
            font_size: 34
          })
          .addText(rows[1], {
            x: 0.2,
            y: 5.2,
            w: 13,
            h: 1,
            align: 'c',
            font_face: 'Montserrat',
            color: 'e9e9e9'
          })
      })
      .catch(err => {
        console.log(err);
        console.log(counter++);
        pptx.addNewSlide()
          .addText(rows[0], {
            x: 0.2,
            y: 4.2,
            w: 13,
            h: 1,
            align: 'c',
            font_face: 'Montserrat',
            color: '76c9d9',
            font_size: 34
          })
          .addText(rows[1], {
            x: 0.2,
            y: 5.2,
            w: 13,
            h: 1,
            align: 'c',
            font_face: 'Montserrat',
            color: 'e9e9e9'
          })
      })
  })
}

Bluebird.resolve().then(() => {
  return csv({noheader:true})
    .fromFile('./data.csv')
    .on('csv', rows => {
      return cetak(rows)
        .delay(100)
        .then(() => {
          if (counter === 251) {
            console.log('saved');
            return pptx.save('test');
          }
        });
    })
    .on('done', (err) => {
      console.log(`done: ${err}`);
    })
});

  
