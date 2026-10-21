// Pengenalan wajah berjalan di browser (face-api.js); server hanya menerima 128 angka wajah.
const MODEL_URL = "https://cdn.jsdelivr.net/npm/@vladmandic/face-api@1.7.13/model/";

async function muatModel() {
  await faceapi.nets.ssdMobilenetv1.loadFromUri(MODEL_URL);
  await faceapi.nets.faceLandmark68Net.loadFromUri(MODEL_URL);
  await faceapi.nets.faceRecognitionNet.loadFromUri(MODEL_URL);
}

async function nyalakanKamera(video) {
  const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "user" } });
  video.srcObject = stream;
  await video.play();
}

async function ambilDescriptor(video) {
  const hasil = await faceapi.detectSingleFace(video).withFaceLandmarks().withFaceDescriptor();
  return hasil ? Array.from(hasil.descriptor) : null;
}

async function kirimJSON(url, isi) {
  const r = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(isi),
  });
  return r.json();
}

function tampilStatus(el, teks, jenis) {
  el.textContent = teks;
  el.className = "status " + (jenis || "");
}

/* =====================================================================
   [TAMBAHAN] Efek tampilan modern — otomatis aktif, kode di atas tidak diubah.
   - Bingkai pemindai (sudut + garis scan) pada kamera
   - Kamera berkedip hijau/merah sesuai hasil absen
   - Animasi muncul pada kotak status
   - Bunyi "bip" + getar HP saat absen berhasil/gagal
   Matikan suara/getar: ubah nilai di bawah menjadi false.
   ===================================================================== */
window.EFEK_SUARA = true;
window.EFEK_GETAR = true;

(function () {
  "use strict";

  // ---- 1) Bingkai pemindai di atas kamera ----
  function pasangPemindai(kamera) {
    if (kamera.querySelector(".pindai")) return;
    const o = document.createElement("div");
    o.className = "pindai";
    o.setAttribute("aria-hidden", "true");
    o.innerHTML =
      '<i class="sudut kiri-atas"></i><i class="sudut kanan-atas"></i>' +
      '<i class="sudut kiri-bawah"></i><i class="sudut kanan-bawah"></i>' +
      '<span class="garis-pindai"></span>';
    kamera.appendChild(o);

    const video = kamera.querySelector("video");
    if (!video) return;
    const hidup = () => kamera.classList.add("aktif");
    const mati = () => kamera.classList.remove("aktif");
    if (!video.paused && video.readyState > 2) hidup();
    video.addEventListener("playing", hidup);
    video.addEventListener("pause", mati);
    video.addEventListener("ended", mati);
  }

  // ---- 2) Bunyi & getar ----
  let ctx = null;
  function bunyi(jenis) {
    if (!window.EFEK_SUARA) return;
    try {
      ctx = ctx || new (window.AudioContext || window.webkitAudioContext)();
      if (ctx.state === "suspended") ctx.resume();
      const o = ctx.createOscillator();
      const g = ctx.createGain();
      const t = ctx.currentTime;
      o.type = "sine";
      o.frequency.value = jenis === "ok" ? 880 : jenis === "galat" ? 220 : 520;
      g.gain.setValueAtTime(0.0001, t);
      g.gain.exponentialRampToValueAtTime(0.12, t + 0.02);
      g.gain.exponentialRampToValueAtTime(0.0001, t + 0.28);
      o.connect(g);
      g.connect(ctx.destination);
      o.start(t);
      o.stop(t + 0.3);
    } catch (e) { /* browser menolak audio: abaikan */ }
  }

  function getar(jenis) {
    if (!window.EFEK_GETAR || !navigator.vibrate) return;
    navigator.vibrate(jenis === "ok" ? 80 : [80, 60, 80]);
  }

  // ---- 3) Reaksi saat kotak status berubah ----
  function jenisStatus(el) {
    return ["ok", "galat", "peringatan"].find((j) => el.classList.contains(j)) || "";
  }

  function kedipKamera(jenis) {
    const kamera = document.querySelector(".kamera");
    if (!kamera || !jenis) return;
    kamera.classList.remove("ok", "galat", "peringatan");
    void kamera.offsetWidth; // paksa animasi mulai ulang
    kamera.classList.add(jenis);
    setTimeout(() => kamera.classList.remove(jenis), 1600);
  }

  function animasiMuncul(el) {
    el.classList.remove("muncul");
    void el.offsetWidth;
    el.classList.add("muncul");
    el.addEventListener("animationend", () => el.classList.remove("muncul"), { once: true });
  }

  const terakhir = new WeakMap();
  function tanggapiStatus(el) {
    const jenis = jenisStatus(el);
    const teks = el.textContent.trim();
    if (!teks) return;
    const kunci = jenis + "|" + teks;
    if (terakhir.get(el) === kunci) return; // hindari pemicu ganda
    terakhir.set(el, kunci);

    animasiMuncul(el);
    if (jenis) {
      kedipKamera(jenis);
      bunyi(jenis);
      getar(jenis);
    }
  }

  function pantauStatus() {
    const pengamat = new MutationObserver((daftar) => {
      const kena = new Set();
      daftar.forEach((m) => {
        const el = m.target.nodeType === 1 ? m.target : m.target.parentElement;
        if (el && el.classList && el.classList.contains("status")) kena.add(el);
      });
      kena.forEach(tanggapiStatus);
    });
    pengamat.observe(document.body, {
      subtree: true, childList: true, characterData: true,
      attributes: true, attributeFilter: ["class"],
    });
  }

  function mulai() {
    document.querySelectorAll(".kamera").forEach(pasangPemindai);
    pantauStatus();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mulai);
  } else {
    mulai();
  }
})();