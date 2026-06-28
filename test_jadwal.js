const url = 'https://script.google.com/macros/s/AKfycbxX9Iw6REGM4JZoMAYzwprf-Tcgx7oZ6uxczgkkURy8okuMQ8GDpKBIo86EypSOSTyVWg/exec';
const body = JSON.stringify({ action: 'getData', args: ['Jadwal'] });

fetch(url, {
  method: 'POST',
  headers: { 'Content-Type': 'text/plain' },
  body: body
})
.then(res => res.json())
.then(data => {
  console.log('JADWAL SUCCESS:', data.success);
  if (data.success) {
    console.log('JADWAL ROWS:', data.rows.length);
    if (data.rows.length > 0) {
      console.log('FIRST JADWAL ROW:', data.rows[0]);
    }
  } else {
    console.log('JADWAL MESSAGE:', data.message);
  }
})
.catch(err => {
  console.error(err);
});
