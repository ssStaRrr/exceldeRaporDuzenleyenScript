const ExcelJS = require("exceljs");
const ChartJS = require("xlsx-chart");


const baskanaGonderilecekRapor = async () => {
  // Orijinal Excel dosyasını oku
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("hamData.xlsx");

  const worksheet = workbook.getWorksheet(1); // okunan excel dosyasının İlk sayfasını al

  // Haric tutmak istediğiniz sütunlar
  const excludeColumns = [
    "#",
    "Takip Kodu",
    "Güncelleme",
    "E-Posta Adresi ",
    "Mesaj",
    "Çalışılan Süre",
    "Bitiş Tarihi",
    "Kurum Sicil Numaranız veya Öğrenci Numaranız",
    "Dahili No / Cep Telefonu",
    "Biriminiz / Bölümünüz",
    "Ödeme Numarası",
  ]; // Haric tutmak istediğiniz sütun adlarını yazın

  // Başlıkları elde et.
  const headers = worksheet.getRow(1).values.slice(1); // İlk satırdaki tüm değerleri array olarak al. ilk değeri atla. çünkü ilk deger boş olarak gelir
  const filteredHeaders = headers.filter(
    (header) => !excludeColumns.includes(header)
  ); // Haric tutulmayan sütunlar

  let data = []; //tüm satırların en son olarak push edildiği data değikeni
  worksheet.eachRow((row, rowNumber) => {
    //sayfadaki her satırı tek tek dön. başlık dışındaki tüm dataları okuyup yeni bir json formatındaki değişkende topla
    if (rowNumber > 1) {
      // İlk satır başlık olduğu için atla
      let rowData = {}; // Yeni excel'imizi olusturan dataların toplanacagı değişkenimiz.
      row.eachCell((cell, colNumber) => {
        //Satırdaki her hücreyi tek tek dön
        const header = worksheet.getRow(1).getCell(colNumber).value; // başlık satırındaki her bir sütun değerini tek tek sırasıyla header değişkeninde tut
        if (!excludeColumns.includes(header)) {
          //bu header değikeni, dışarda tuttuğumuz header değişkeni ile eşleşmiyorsa bu header sütunu altındaki hücredeki değikeni tutabilirsin
          rowData[header] = cell.value;
        }
      });

      // Tarih formatını düzenle
      if (rowData["Tarih"]) {
        rowData["Tarih"] = new Date(rowData["Tarih"]).toLocaleDateString(); // İstediğiniz formata çevirin
      }

      data.push(rowData);
    }
  });

  // Yeni 1 adet excel + 2 adet sayfa oluştur
  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("MSGSÜ Tüm Birimler");
  const newWorksheet2 = newWorkbook.addWorksheet("Bilgi İşlem");

  // Başlıkları ekle - sayfa 1 ve sayfa2 ye
  const headerRow = newWorksheet.addRow(filteredHeaders);
  const headerRow2 = newWorksheet2.addRow(filteredHeaders);

  filteredHeaders.forEach((header, index) => {
    //birinci ve ikinci sayda
    const cell = headerRow.getCell(index + 1);
    const cell2 = headerRow2.getCell(index + 1);

    cell2.font = { bold: true };
    cell2.alignment = {
      horizontal: "left",
      vertical: "middle", // Dikeyde ortala
      wrapText: true, // Metni sarmala
    };
    cell2.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };

    cell.font = { bold: true };
    cell.alignment = {
      horizontal: "left",
      vertical: "middle", // Dikeyde ortala
      wrapText: true, // Metni sarmala
    };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });

  let addThinToCell = () => {};

  // grafiğe işlenececek datalar
  let cozuldu = 0;
  let cevaplandi = 0;
  let yeni = 0;
  let islemGoruyor = 0;
  let cevapBekliyor = 0;
  let beklemede = 0;
  // Veriyi ekle - ilk sayfa için
  data.forEach((row, index) => {
    const newRow = newWorksheet.addRow(Object.values(row)); //1. sayfaya, data oldugu gibi ekle.

    //İkinci Sayfa için Kategori "Bilgi İşlem Daire Başkanlığı" olan değerleri tut ve 2. sayfaya at
    const categoryValue = Object.values(row)[2];
    if (
      categoryValue.includes("Bilgi İşlem Dairesi Başkanlığı") ||
      categoryValue.includes("EBYS ve E-imza İşlemleri")
    ) {
      let talepDurumu = row["Talep Durumu"];
      if (talepDurumu === "Çözüldü") {
        cozuldu += 1;
      } else if (talepDurumu === "Cevaplandı") {
        cevaplandi += 1;
      } else if (talepDurumu === "Yeni") {
        yeni += 1;
      } else if (talepDurumu === "İşlem Görüyor") {
        islemGoruyor += 1;
      } else if (talepDurumu === "Cevap Bekliyor") {
        cevapBekliyor += 1;
      } else if (talepDurumu === "Beklemede") {
        beklemede += 1;
      }

      //sütun 3'te içinde bilgi işlem kelimesi geçen her satırı addRow ile ekle
      const newRowInSheet2 = newWorksheet2.addRow(Object.values(row));

      newRowInSheet2.height = 35;
      // Her hücreye kenarlık ekle
      newRowInSheet2.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });
    }

    // Satır yüksekliği ayarla (örneğin, 20 birim yüksekliği)
    newRow.height = 35;

    // Her hücreye kenarlık ekle
    newRow.eachCell((cell) => {
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
  });

  // Sütun genişliklerini ayarla - 1.sayfa için
  newWorksheet.columns = [
    { width: 15 }, // Talep No sütunu genişliği
    { width: 20 }, // Geliş Tarihi sütunu genişliği
    { width: 35 }, // Bitiş Tarihi sütunu genişliği
    { width: 12 }, // Gönderen Kişi sütunu genişliği
    { width: 12 }, // Gelen Birim sütunu genişliği
    { width: 40 }, // Sorun sütunu genişliği
    { width: 30 }, // Bilgisayar Özellikleri sütunu genişliği
  ];

  // Sütun genişliklerini ayarla - 2.sayfa için
  newWorksheet2.columns = [
    { width: 15 }, // Talep No sütunu genişliği
    { width: 25 }, // Geliş Tarihi sütunu genişliği
    { width: 45 }, // Bitiş Tarihi sütunu genişliği
    { width: 12 }, // Gönderen Kişi sütunu genişliği
    { width: 12 }, // Gelen Birim sütunu genişliği
    { width: 40 }, // Sorun sütunu genişliği
    { width: 30 }, // Bilgisayar Özellikleri sütunu genişliği
  ];

  const row2 = newWorksheet2.getRow(2);
  const row3 = newWorksheet2.getRow(3);
  const row4 = newWorksheet2.getRow(4);
  const row5 = newWorksheet2.getRow(5);
  const row6 = newWorksheet2.getRow(6);
  const row7 = newWorksheet2.getRow(7);
  const row8 = newWorksheet2.getRow(8);

  row2.getCell(9).value = "İŞLEM SONUCU";
  newWorksheet2.getColumn(9).width = 18;
  row2.getCell(10).value = "ADET";

  row3.getCell(9).value = "Çözüldü";
  row3.getCell(10).value = cozuldu;
  row4.getCell(9).value = "Cevaplandı";
  row4.getCell(10).value = cevaplandi;
  row5.getCell(9).value = "Yeni";
  row5.getCell(10).value = yeni;
  row6.getCell(9).value = "İşlem Görüyor";
  row6.getCell(10).value = islemGoruyor;
  row7.getCell(9).value = "Cevap Bekliyor";
  row7.getCell(10).value = cevapBekliyor;
  row8.getCell(9).value = "Beklemede";
  row8.getCell(10).value = beklemede;

  row2.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  row3.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  row4.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  row5.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  row6.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  row7.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });
  row8.eachCell((cell) => {
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
  });

  // const chartOptions = {
  //   title: "Taleplerin Değerlendirilmesi",
  //   chart: "pie",
  //   series: [
  //     {
  //       name: "Talepler",
  //       labels: ["Çözüldü", "Cevaplandı", "Yeni", "İşlem Görüyor", "Cevap Bekliyor", "Beklemede" ],
  //       values: [row3.getCell(10).value, row4.getCell(10).value,row5.getCell(10).value,row6.getCell(10).value,row7.getCell(10).value,row8.getCell(10).value],
  //     },
  //   ],
  // };

  // const chart = new ChartJS.Chart(chartOptions);
  // chart.render(newWorksheet2, "B15"); // Grafiği belirtilen hücreye yerleştir

  // Yeni Excel dosyasına yaz
  await newWorkbook.xlsx.writeFile(
    "EYLÜL 2024 MSGSU-BİDB Talep Yönetim Sistemi Raporu.xlsx"
  );

  console.log("BİDB Daire Başkanına Gönderilecek Rapor Oluşturuldu");
};

module.exports = {
  baskanaGonderilecekRapor,
};
