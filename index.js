
const ExcelJS = require('exceljs');
const path = require('path');

// Natural sorting için bir yardımcı fonksiyon
function naturalSort(a, b) {
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

// Excel dosyasını oku
const workbook = new ExcelJS.Workbook();
const filePath = path.join(__dirname, 'kesim-adetleri.xlsx');
function turkceKarakterleriDegistir(metin) {
    const karakterler = {
        'İ': 'I',
        'ı': 'i',
        'Ş': 'S',
        'ş': 's',
        'Ç': 'C',
        'ç': 'c',
        'Ğ': 'G',
        'ğ': 'g',
        'Ö': 'O',
        'ö': 'o',
        'Ü': 'U',
        'ü': 'u'
    };

    return metin.split('').map(karakter => karakterler[karakter] || karakter).join('');
}


// let renk = turkceKarakterleriDegistir(renkk);




workbook.xlsx.readFile(filePath).then(() => {
    const worksheet = workbook.getWorksheet(1); // İlk sayfayı seçin
    const totals = {};

    // Her satırı oku
    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Başlık satırını atla

        const model = row.getCell(1).value;  // Model sütunu (1. sütun)
        const siparis = row.getCell(2).value; // Sipariş no sütunu (2. sütun)
        const kesimAdeti = row.getCell(3).value; // Kesim Adeti sütunu (3. sütun)
        const renkTurkce = row.getCell(4).value; // VARYANT sütunu (4. sütun)
        const renk = turkceKarakterleriDegistir(renkTurkce);
        // Toplamları hesapla
        if (!totals[renk]) totals[renk] = {};
        if (!totals[renk][model]) totals[renk][model] = {};
        if (!totals[renk][model][siparis]) totals[renk][model][siparis] = 0;

        totals[renk][model][siparis] += kesimAdeti;
        console.log(totals);
    });

    // Model ve sipariş numaralarını doğal sıralama ile sıralı hale getir
    const sortedData = Object.keys(totals).sort().reduce((acc, renk) => {
        acc[renk] = Object.keys(totals[renk]).sort(naturalSort).reduce((modelAcc, model) => {
            modelAcc[model] = Object.keys(totals[renk][model]).sort(naturalSort).reduce((siparisAcc, siparis) => {
                siparisAcc[siparis] = totals[renk][model][siparis];
                return siparisAcc;
            }, {});
            return modelAcc;
        }, {});
        return acc;
    }, {});

    // Yeni bir workbook ve worksheet oluştur
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Toplam Kesim Adetleri');

    // Başlık satırı ekle (sizin belirttiğiniz sıraya göre)
    newWorksheet.addRow(['Model', 'Sipariş no', 'Kesim adedi', 'VARYANT']);

    // Sıralanmış verileri yeni worksheet'e ekle
    Object.keys(sortedData).forEach(renk => {
        Object.keys(sortedData[renk]).forEach(model => {
            Object.keys(sortedData[renk][model]).forEach(siparis => {
                const kesimAdeti = sortedData[renk][model][siparis];
                newWorksheet.addRow([model, siparis, kesimAdeti, renk]);
            });
        });
    });

    // Yeni dosyayı kaydet
    const outputPath = path.join(__dirname, 'toplam-kesim-adetleri.xlsx');
    return newWorkbook.xlsx.writeFile(outputPath);
}).then(() => {
    console.log('Yeni Excel dosyası oluşturuldu: toplam-kesim-adetleri.xlsx');
}).catch(err => {
    console.error('Bir hata oluştu:', err);
});
