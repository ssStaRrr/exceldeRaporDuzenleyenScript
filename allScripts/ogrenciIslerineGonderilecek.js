const ExcelJS = require("exceljs");
const fs = require("fs");

const ogrenciIslerineGonderilecek = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("hamData.xlsx");

  const worksheet = workbook.getWorksheet(1); // İlk sayfayı al

  // Veriyi işlemek
  const excludeColumns = [
    "#",
    "Takip Kodu",
    "E-Posta Adresi ",
    //"Kategori",
    "Mesaj",
    "Çalışılan Süre",
    "Bitiş Tarihi",
    "Kurum Sicil Numaranız veya Öğrenci Numaranız",
    "Dahili No / Cep Telefonu",
    "Ödeme Numarası",
  ];

  let data = [];

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // İlk satır başlık olduğu için atla
      let rowData = {};
      row.eachCell((cell, colNumber) => {
        const header = worksheet.getRow(1).getCell(colNumber).value;
        if (!excludeColumns.includes(header)) {
          rowData[header] = cell.value;
        }
      });
      if (
        rowData["Kategori"] &&
        (rowData["Kategori"].includes("Öğrenci İşleri Daire Başkanlığı") ||
          rowData["Kategori"].includes("Öğrenci İşleri Dairesi Başkanlığı"))
      ) {
        data.push(rowData);
      }
    }
  });
  // Yeni bir sayfa oluştur ve verileri yaz
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("TYS");

  // Üste boş bir satır ekle
  newWorksheet.addRow([]);

  // Başlıkları ekle
  const headers = [
    "",
    "Talep Sahibi ",
    "Biriminiz / Bölümünüz",
    "Talep Konusu ",
    "Önem Derecesi ",
    "Sahibi",
    "Tarih",
    "Talep Durumu",
    "Güncelleme",
    "Talep Çözümlenmediyse Nedeni",
  ];

  // Başlıkları ekle
  const newHeaders = [
    "",
    "Talep Sahibi",
    "Talep Sahibinin Birimi (Personel veya öğrenciyse hangi birime bağlı olduğu)",
    "Talep Konusu",
    "Önem Derecesi",
    "Taleple İlgilenen Personelin Adı",
    "Talebin Birime Ulaşma Tarihi",
    "Talebin Durumu",
    "Talebin Sonuçlanma Tarihi",
    "Talep Çözümlenmediyse Nedeni",
  ];

  newWorksheet.addRow(headers);

  // A1'den J1'e kadar hücreleri birleştir ve bir yazı ekle
  newWorksheet.mergeCells("A1:J1");
  const mergedCell = newWorksheet.getCell("A1");
  mergedCell.value = "Talebin Ulaştığı Birim: ÖĞRENCİ İŞLERİ DAİRE BAŞKANLIĞI"; // İstenilen metin
  mergedCell.font = { bold: true, size: 14 }; // Fontu kalın ve büyük yap
  mergedCell.alignment = { horizontal: "left", vertical: "middle" }; // Ortala
  newWorksheet.getRow(1).height = 50; // Yüksekliği artır

  // Başlık satırının yüksekliğini artır
  newWorksheet.getRow(2).height = 55; // 45 olarak ayarlandı, bu değeri artırabilirsiniz

  // Kenarlık ekleme fonksiyonu
  const setBorders = (cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  };

  // Başlıkları kalın ve ortalanmış yap
  headers.forEach((header, index) => {
    const cell = newWorksheet.getRow(2).getCell(index + 1);
    cell.font = { bold: true };
    cell.alignment = {
      horizontal: "center",
      vertical: "middle", // Dikeyde ortala
      wrapText: true // Metni sarmala
    };
    setBorders(cell);
  });

  // Veriyi ekle
  let cozuldu = 0
  let cozulmedi = 0
  let cevapBekliyor = 0
  data.forEach((row, index) => {
    const rowData = headers.map((header) => {
      
      if(row[header]==="Yeni") {
        row[header] = "Cevap Bekliyor"
      }
      if(row[header] === "Cevaplandı") {
        row[header] = "Cevap Bekliyor"
      }
      if(row[header] === "Cevap Bekliyor") {
        cevapBekliyor+=1
      }
      if(row[header] === "Çözüldü") {
        cozuldu+=1
      }
      return row[header] || ""
    }); // Verilerde boş alanları düzelt
      
  
    // İlk sütuna artan sıra numarasını ekle
    rowData[0] = index + 1;
    const newRow = newWorksheet.addRow(rowData); // Verileri ekle
    newRow.eachCell((cell) => {
      setBorders(cell); // Her hücreye kenarlık ekle
    });
  });

  let i = 0;
  newWorksheet.getRow(2).eachCell((cell) => {
    cell.value = newHeaders[i];
    i++;
  });

  // Sütun genişliklerini ayarla
  newWorksheet.columns = [
    { width: 10 },
    { width: 25 },
    { width: 35 },
    { width: 40 },
    { width: 15 },
    { width: 35 },
    { width: 16 },
    { width: 15 },
    { width: 15 },
    { width: 30 },
  ];

  // En alta iki boş satır ekle
  newWorksheet.addRow([]);

  const secondLastRow = newWorksheet.addRow([]); // İkinci boş satır
  const thirdLastRow = newWorksheet.addRow([])
  const fourthLastRow = newWorksheet.addRow([])
  const fifthLastRow = newWorksheet.addRow([])

  secondLastRow.getCell(2).value = "NOT: Talebin Durumuna yalnızca yan taraftaki tabloda yer alan ifadelerden biri yazılmalıdır. "
  secondLastRow.getCell(8).value = "Talep Durumu:"
  secondLastRow.getCell(9).value = "Adet:"
  secondLastRow.font = { bold: true, size:14 }; 
  secondLastRow.alignment = { horizontal: 'left' }; // Yazıyı sola hizala
  thirdLastRow.getCell(8).value = "Çözüldü"
  thirdLastRow.getCell(9).value = cozuldu
  fourthLastRow.getCell(8).value = "Çözümlenmedi"
  fourthLastRow.getCell(9).value = cozulmedi
  fifthLastRow.getCell(8).value = "Cevap Bekliyor"
  fifthLastRow.getCell(9).value = cevapBekliyor
  setBorders(secondLastRow.getCell(8))
  setBorders(secondLastRow.getCell(9))
  setBorders(thirdLastRow.getCell(8))
  setBorders(thirdLastRow.getCell(9))
  setBorders(fourthLastRow.getCell(8))
  setBorders(fourthLastRow.getCell(9))
  setBorders(fifthLastRow.getCell(8))
  setBorders(fifthLastRow.getCell(9))
  // Yeni dosyayı kaydet
  await newWorkbook.xlsx.writeFile("Eylül 2024 Öğrenci İşleri TYS.xlsx");
};

console.log("Öğrenci İlerine Gönderilecek Rapor Oluşturuldu")

module.exports = {
  ogrenciIslerineGonderilecek,
};
