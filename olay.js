/* Canlı adres: bu satır, adres.txt ve manifest.xml içindeki adres aynı olmalı. */
var HLA_HOST = "https://hlalper.github.io/hba-outlook-font-ctb";
var HLA_PAKET_JSON = HLA_HOST + "/paket.json";
var HLA_LINK_ADI = "HBA_PlotCfgs_Fonts.zip";
var HLA_LINK_ADRESI = "https://hbams-my.sharepoint.com/personal/lutfu_alper_hba_com/_layouts/15/download.aspx?share=IQC3UI6UU1iFSoj3wzX6IHaTASL6HGvxE6_MxkP7IM4PC48";
var HLA_SINIR_BAYT = 25 * 1024 * 1024;
var HLA_IZLER = [
  "çizgi kalemi",
  "plot style",
  "plotstyle",
  "e-transmit",
  "etransmit",
  "autocad",
  "çizim",
  "pafta",
  "font",
  "ctb",
  "dwg",
  "dxf",
  "shx",
  "xref",
  "proje",
  "project",
  "drawing"
];

function hlaNorm(metin) {
  return String(metin || "").replace(/\u0130/g, "i").toLowerCase();
}

function hlaIzVar(metin) {
  var bak = hlaNorm(metin);
  if (!bak) return false;
  for (var i = 0; i < HLA_IZLER.length; i++) {
    if (bak.indexOf(hlaNorm(HLA_IZLER[i])) !== -1) return true;
  }
  return false;
}

function hlaBaytYazi(bayt) {
  var n = Number(bayt) || 0;
  if (n < 1024) return n + " B";
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
  return (n / (1024 * 1024)).toFixed(1) + " MB";
}

function hlaTarihYazi(iso) {
  if (!iso) return "";
  var d = new Date(iso);
  if (isNaN(d.getTime())) return String(iso);
  function iki(n) { return n < 10 ? "0" + n : String(n); }
  return d.getDate() + "." + (d.getMonth() + 1) + "." + d.getFullYear() + " " + iki(d.getHours()) + ":" + iki(d.getMinutes());
}

function hlaKisa(metin) {
  var s = String(metin || "");
  return s.length > 140 ? s.slice(0, 137) + "..." : s;
}

function hlaPaketOku(bitti) {
  var url = HLA_PAKET_JSON + "?t=" + Date.now();
  if (typeof fetch === "function") {
    fetch(url, { cache: "no-store" })
      .then(function (yanit) {
        if (!yanit.ok) throw new Error("HTTP " + yanit.status);
        return yanit.json();
      })
      .then(function (paket) { bitti(paket, null); })
      .catch(function (hata) { bitti(null, hata); });
    return;
  }
  try {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", url, true);
    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;
      if (xhr.status < 200 || xhr.status >= 300) {
        bitti(null, new Error("HTTP " + xhr.status));
        return;
      }
      try {
        bitti(JSON.parse(xhr.responseText), null);
      } catch (e) {
        bitti(null, e);
      }
    };
    xhr.onerror = function () { bitti(null, new Error("ağ")); };
    xhr.send();
  } catch (e) {
    bitti(null, e);
  }
}

function hlaHtmlKacis(metin) {
  return String(metin || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function hlaDuzYazi(html) {
  return String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">");
}

function hlaLinkJeton(adres) {
  var m = String(adres || "").match(/[?&]share=([^&#]+)/i);
  if (m) return decodeURIComponent(m[1]);
  m = String(adres || "").match(/\/([A-Za-z0-9_-]{20,})(?:\?|$)/);
  return m ? m[1] : "";
}

function hlaLinkVar(html, paket) {
  var bak = String(html || "");
  if (!bak) return false;
  if (paket.dosyaAdi && bak.indexOf(paket.dosyaAdi) !== -1) return true;
  if (paket.adres && bak.indexOf(paket.adres) !== -1) return true;
  var jeton = hlaLinkJeton(paket.adres);
  return !!(jeton && bak.indexOf(jeton) !== -1);
}

function hlaLinkParcasi(paket) {
  return '<p><a href="' + hlaHtmlKacis(paket.adres) + '">' + hlaHtmlKacis(paket.dosyaAdi) + "</a></p>";
}

function hlaSonaEkle(html, parca) {
  var n = String(html || "").toLowerCase().lastIndexOf("</body>");
  if (n < 0) return String(html || "") + parca;
  return html.slice(0, n) + parca + html.slice(n);
}

function hlaLinkiYaz(item, paket, bitti) {
  if (!item.body || !item.body.getAsync || !item.body.setAsync) {
    bitti({ tamam: false, mesaj: "Bu ekranda mektup yazısı yok." });
    return;
  }
  item.body.getAsync(Office.CoercionType.Html, function (okuma) {
    if (okuma.status !== Office.AsyncResultStatus.Succeeded) {
      bitti({ tamam: false, mesaj: "Mektup yazısı okunamadı." });
      return;
    }
    if (hlaLinkVar(okuma.value, paket)) {
      bitti({ tamam: true, mesaj: "Font/CTB linki zaten var.", zaten: true });
      return;
    }
    item.body.setAsync(hlaSonaEkle(okuma.value, hlaLinkParcasi(paket)), { coercionType: Office.CoercionType.Html }, function (yazma) {
      if (yazma.status !== Office.AsyncResultStatus.Succeeded) {
        var detay = yazma.error && yazma.error.message ? yazma.error.message : "yazılamadı";
        bitti({ tamam: false, mesaj: hlaKisa("Link eklenemedi: " + detay) });
        return;
      }
      bitti({ tamam: true, mesaj: "Font/CTB linki eklendi." });
    });
  });
}

function hlaYerelPaket() {
  return { dosyaAdi: HLA_LINK_ADI, adres: HLA_LINK_ADRESI };
}

function hlaPaketiEkle(bitti) {
  var item = Office.context.mailbox.item;
  if (!item) {
    bitti({ tamam: false, mesaj: "Bu ekranda mektup yok." });
    return;
  }
  hlaLinkiYaz(item, hlaYerelPaket(), bitti);
}

function hlaBildir(sonuc) {
  try {
    var bilgi = sonuc.tamam
      ? Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage
      : Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage;
    var govde = { type: bilgi, message: hlaKisa(sonuc.mesaj) };
    if (sonuc.tamam) {
      govde.icon = "Icon.16x16";
      govde.persistent = false;
    }
    Office.context.mailbox.item.notificationMessages.replaceAsync("hlaFontCtb", govde);
  } catch (e) {}
}

function ekleKomutu(event) {
  var kapandi = false;
  function kapat() {
    if (kapandi) return;
    kapandi = true;
    try {
      if (event && typeof event.completed === "function") event.completed();
    } catch (e) {}
  }
  try {
    if (typeof setTimeout === "function") setTimeout(kapat, 3000);
    var item = Office.context.mailbox.item;
    var parca = hlaLinkParcasi(hlaYerelPaket());
    var opts = { coercionType: Office.CoercionType.Html };
    if (item && item.body && item.body.prependAsync) {
      item.body.prependAsync(parca, opts, kapat);
      return;
    }
    if (item && item.body && item.body.setSelectedDataAsync) {
      item.body.setSelectedDataAsync(parca, opts, kapat);
      return;
    }
  } catch (e) {}
  kapat();
}

function hlaHatirlat(bitir, mesaj) {
  var opts = {
    allowEvent: false,
    errorMessage: mesaj,
    cancelLabel: "Linki ekle",
    commandId: "ekleKomutu"
  };
  if (Office.MailboxEnums && Office.MailboxEnums.SendModeOverride) {
    opts.sendModeOverride = Office.MailboxEnums.SendModeOverride.PromptUser;
  }
  bitir(opts);
}

function onMessageSendHandler(event) {
  var bitti = false;
  function bitir(opts) {
    if (bitti) return;
    bitti = true;
    try { event.completed(opts); } catch (e) {}
  }
  var paket = hlaYerelPaket();
  var item = null;
  try { item = Office.context.mailbox.item; } catch (e0) { item = null; }
  if (!item || !item.subject || !item.subject.getAsync || !item.body || !item.body.getAsync) {
    bitir({ allowEvent: true });
    return;
  }
  item.subject.getAsync(function (konuSonuc) {
    var konu = "";
    try {
      if (konuSonuc && konuSonuc.status === Office.AsyncResultStatus.Succeeded) konu = konuSonuc.value || "";
    } catch (e1) {}
    item.body.getAsync(Office.CoercionType.Text, function (yaziSonuc) {
      try {
        var yazi = "";
        if (yaziSonuc && yaziSonuc.status === Office.AsyncResultStatus.Succeeded) yazi = yaziSonuc.value || "";
        var birlesik = konu + "\n" + yazi;
        if (!hlaIzVar(birlesik) || hlaLinkVar(birlesik, paket)) bitir({ allowEvent: true });
        else hlaHatirlat(bitir, "Font/CTB linkini eklemedin.");
      } catch (e2) {
        bitir({ allowEvent: true });
      }
    });
  });
}

function hlaBagla() {
  if (typeof Office === "undefined" || !Office.actions || !Office.actions.associate) return;
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("ekleKomutu", ekleKomutu);
}

hlaBagla();
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function () { hlaBagla(); });
}
