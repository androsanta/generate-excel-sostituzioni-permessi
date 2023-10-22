
const currentYear = new Date().getFullYear();
document.getElementById('annoscolastico').placeholder = String(currentYear);

function onSubmit() {
    const teachers = document.getElementById('listadocenti').value;
    const year = document.getElementById('annoscolastico').value;
    generateExcel(teachers, year ? Number(year) : undefined)
        .then(() => {
            console.log('all good!')
        })
        .catch(e => {
            console.error('Error generating file', e);
        });
}