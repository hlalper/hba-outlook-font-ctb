function hlaGoster(el, acik) {
  if (!el) return;
  if (acik) el.removeAttribute("hidden");
  else el.setAttribute("hidden", "");
}

function hlaAyarCiz(paket) {
  document.getElementById("adres").textContent = paket.adres || "";
  document.getElementById("dosya").textContent = paket.dosyaAdi || "";
  document.getElementById("boyut").textContent = hlaBaytYazi(paket.bayt);
  document.getElementById("tarih").textContent = hlaTarihYazi(paket.tarih);
  var liste = document.getElementById("liste");
  liste.textContent = "";
  var dosyalar = paket.dosyalar || [];
  if (!dosyalar.length) {
    var bos = document.createElement("li");
    bos.textContent = "Zip boş.";
    liste.appendChild(bos);
  } else {
    for (var i = 0; i < dosyalar.length; i++) {
      var li = document.createElement("li");
      li.textContent = dosyalar[i];
      liste.appendChild(li);
    }
  }
  var uyari = document.getElementById("uyari");
  if (Number(paket.bayt) > HLA_SINIR_BAYT) {
    uyari.textContent = "Bu zip 25 MB sınırının üstünde. Outlook tek ek olarak kabul etmeyebilir.";
    hlaGoster(uyari, true);
  } else {
    uyari.textContent = "";
    hlaGoster(uyari, false);
  }
  hlaGoster(document.getElementById("kart"), true);
  hlaGoster(document.getElementById("ic"), true);
  document.getElementById("durum").textContent = "";
}

function hlaAyarYukle(outlookta) {
  hlaPaketOku(function (paket, hata) {
    if (hata || !paket) {
      document.getElementById("durum").textContent = "Pakete ulaşılamadı. Adres: " + HLA_PAKET_JSON;
      return;
    }
    hlaAyarCiz(paket);
    if (outlookta) hlaGoster(document.getElementById("ekle"), true);
  });
}

Office.onReady(function (info) {
  var outlookta = info.host === Office.HostType.Outlook && Office.context.mailbox && Office.context.mailbox.item;
  document.getElementById("ekle").addEventListener("click", function () {
    var sonuc = document.getElementById("sonuc");
    sonuc.textContent = "Ekleniyor.";
    hlaPaketiEkle(function (cevap) {
      sonuc.textContent = cevap.mesaj;
      hlaBildir(cevap);
    });
  });
  hlaAyarYukle(!!outlookta);
  if (outlookta && /(?:\?|&)ekle=1(?:&|$)/.test(window.location.search)) {
    document.getElementById("sonuc").textContent = "Ekleniyor.";
    hlaPaketiEkle(function (cevap) {
      document.getElementById("sonuc").textContent = cevap.mesaj;
      hlaBildir(cevap);
    });
  }
});
