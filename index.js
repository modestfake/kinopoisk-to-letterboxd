let readXlsxFile = require('read-excel-file/node')
let fs = require('fs')

function formatTitle (title) {
  title = title.toString()

  if (title.includes('"')) {
    title = title.replace('"', '\\"')
  }

  if (title.includes(',')) {
    title = `"${title}"`
  }

  return title
}

function formatYear (year) {
  if (typeof year !== 'number') {
    year = Number(year.split(' ')[0])
  }

  return year
}

let csvHeading = 'Title,Year,Rating10'

async function run () {
  try {
    let filename = process.argv[2]

    if (!filename) throw new Error('XLS file required')

    let data = await readXlsxFile(filename)

    let films = data
      .slice(1)
      .map(film => ({
        title: film[1] || film[0],
        year: film[2],
        rating: film[3]
      }))

    let csvData = films.map(film => (
      `${formatTitle(film.title)},${formatYear(film.year)},${film.rating}`
    )).join('\n')

    let csv = [csvHeading, csvData].join('\n')

    fs.writeFileSync('result.csv', csv)
  } catch (error) {
    console.error(error)
  }
}

run()
