<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Input Nilai Mahasiswa</title>
  <link href="https://fonts.googleapis.com/css2?family=Press+Start+2P&display=swap" rel="stylesheet">
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    :root {
      --bg-main: #0f0f1a;
      --text-main: #00ffff;
      --text-secondary: #c0c0c0;
      --accent: #ff00c8;
      --hover: #7a00ff;
      --border: #1f1f3d;
      --surface: #14141f;
    }
    * {
      box-sizing: border-box;
      font-family: 'Press Start 2P', monospace;
      transition: all 0.3s ease;
    }
    body {
      margin: 0;
      padding: 0;
      background-color: var(--bg-main);
      color: var(--text-main);
    }
    .container {
      max-width: 1100px;
      margin: 3rem auto;
      padding: 2rem;
      background-color: var(--surface);
      border: 1px solid var(--border);
      border-radius: 10px;
      box-shadow: 0 0 20px #00ffff50;
      animation: fadeIn 1s ease-in-out;
    }
    h1, h3 {
      text-align: center;
      color: var(--accent);
      text-shadow: 0 0 10px var(--accent);
    }
    input[type="number"], input[type="text"], input[type="file"] {
      background-color: #1f1f3d;
      color: var(--text-main);
      border: 1px solid var(--accent);
      border-radius: 6px;
      padding: 0.5rem;
      margin: 0.3rem 0;
      width: 100%;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 1rem;
    }
    th, td {
      padding: 0.8rem;
      border: 1px solid var(--border);
      text-align: center;
    }
    thead th {
      background-color: var(--border);
      color: var(--accent);
    }
    button {
      background: linear-gradient(45deg, var(--accent), var(--hover));
      border: none;
      color: white;
      padding: 0.6rem 1.2rem;
      border-radius: 8px;
      margin: 0.4rem;
      cursor: pointer;
      box-shadow: 0 0 10px var(--accent), 0 0 20px var(--hover);
    }
    button:hover {
      transform: scale(1.05);
      box-shadow: 0 0 20px var(--hover), 0 0 30px var(--accent);
    }
    .btn-hapus {
      background-color: var(--border);
      color: var(--text-secondary);
      border-radius: 6px;
    }
    .btn-hapus:hover {
      background-color: var(--accent);
      color: white;
    }
    #hasil {
      margin-top: 2rem;
    }
    .toggle-darkmode {
      position: fixed;
      top: 1rem;
      right: 1rem;
      background-color: var(--surface);
      color: var(--text-main);
      border: 1px solid var(--accent);
      padding: 0.5rem 1rem;
      border-radius: 6px;
    }
    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(-10px); }
      to { opacity: 1; transform: translateY(0); }
    }
    @media (max-width: 768px) {
      h1, h3 {
        font-size: 1.3rem;
      }
      button {
        font-size: 0.9rem;
        padding: 0.4rem 0.8rem;
      }
    }
  </style>
</head>
<body>
  <button class="toggle-darkmode" onclick="document.body.classList.toggle('dark')">🌓 Mode Gelap</button>
  <div class="container">
    <h1>Input Nilai Mahasiswa</h1>
    <input type="number" id="jumlahPertemuanInput" placeholder="Jumlah Pertemuan" min="1" />
    <button onclick="buatTabelAwal()">Buat Tabel</button>
    <button onclick="tambahPertemuan()">Tambah Pertemuan</button>
    <button onclick="hapusPertemuan()">Hapus Pertemuan</button>
    <table id="tabelMahasiswa">
      <thead>
        <tr>
          <th>Nama</th>
          <th>Aksi</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
    <button onclick="tambahBaris()">➕ Tambah Mahasiswa</button>
    <button onclick="hitungSemua()">🧮 Hitung Semua</button>
    <input type="file" id="excelInput" accept=".xlsx, .xls" onchange="importDariExcel(event)" />
    <div id="hasil">
      <h3>Hasil Kalkulasi</h3>
      <table id="tabelHasil">
        <thead>
          <tr>
            <th>Nama</th>
            <th>Total</th>
            <th>Rata-rata</th>
            <th>Nilai Huruf</th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>
  </div>

  <div style="text-align: center; margin-bottom: 2rem;">
    <button onclick="exportKeExcel()">📁 Export ke Excel</button>
  </div>
  <script>
    let jumlahPertemuan = 0;
function buatTabelAwal() {
  const inputJumlah = parseInt(
    document.getElementById("jumlahPertemuanInput").value
  );
  if (isNaN(inputJumlah) || inputJumlah < 1) {
    alert("Masukkan jumlah pertemuan minimal 1.");
    return;
  }
  jumlahPertemuan = inputJumlah;
  const theadRow = document.querySelector("#tabelMahasiswa thead tr");
  theadRow.innerHTML = "<th>Nama</th>";
  for (let i = 1; i <= jumlahPertemuan; i++) {
    const th = document.createElement("th");
    th.textContent = `P${i}`;
    theadRow.appendChild(th);
  }
  theadRow.innerHTML += "<th>Aksi</th>";
  document.querySelector("#tabelMahasiswa tbody").innerHTML = "";
  tambahBaris();
}
function tambahPertemuan() {
  jumlahPertemuan++;
  const theadRow = document.querySelector("#tabelMahasiswa thead tr");
  const thBaru = document.createElement("th");
  thBaru.textContent = `P${jumlahPertemuan}`;
  theadRow.insertBefore(thBaru, theadRow.lastElementChild);
  document.querySelectorAll("#tabelMahasiswa tbody tr").forEach((tr) => {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "number";
    input.className = "nilai";
    td.appendChild(input);
    tr.insertBefore(td, tr.lastElementChild);
  });
}
function hapusPertemuan() {
  if (jumlahPertemuan <= 1) {
    alert("Minimal harus ada 1 pertemuan.");
    return;
  }
  jumlahPertemuan--;
  const theadRow = document.querySelector("#tabelMahasiswa thead tr");
  theadRow.removeChild(theadRow.children[theadRow.children.length - 2]);
  document.querySelectorAll("#tabelMahasiswa tbody tr").forEach((tr) => {
    tr.removeChild(tr.children[tr.children.length - 2]);
  });
}
function tambahBaris() {
  const tbody = document.querySelector("#tabelMahasiswa tbody");
  const tr = document.createElement("tr");
  const tdNama = document.createElement("td");
  const inputNama = document.createElement("input");
  inputNama.type = "text";
  inputNama.className = "nama";
  inputNama.placeholder = "Nama mahasiswa";
  tdNama.appendChild(inputNama);
  tr.appendChild(tdNama);
  for (let i = 0; i < jumlahPertemuan; i++) {
    const td = document.createElement("td");
    const input = document.createElement("input");
    input.type = "number";
    input.className = "nilai";
    td.appendChild(input);
    tr.appendChild(td);
  }
  const tdAksi = document.createElement("td");
  const btnHapus = document.createElement("button");
  btnHapus.textContent = "Hapus";
  btnHapus.className = "btn-hapus";
  btnHapus.onclick = () => tr.remove();
  tdAksi.appendChild(btnHapus);
  tr.appendChild(tdAksi);
  tbody.appendChild(tr);
  aktifkanAutoTab();
}
function hitungSemua() {
  const rows = document.querySelectorAll("#tabelMahasiswa tbody tr");
  const hasilBody = document.querySelector("#tabelHasil tbody");
  hasilBody.innerHTML = "";
  rows.forEach((row) => {
    const nama = row.querySelector(".nama").value.trim();
    const nilaiInputs = row.querySelectorAll(".nilai");
    let total = 0;
    let valid = true;
    nilaiInputs.forEach((input) => {
      const nilai = parseFloat(input.value);
      if (isNaN(nilai)) {
        input.style.border = "1px solid red";
        valid = false;
      } else {
        input.style.border = "";
        total += nilai;
      }
    });
    if (nama === "" || !valid) return;
    const rata = total / jumlahPertemuan;
    const huruf = konversiHuruf(rata);
    const hasilRow = document.createElement("tr");
    hasilRow.innerHTML = `
          <td>${nama}</td>
          <td>${total.toFixed(2)}</td>
          <td>${rata.toFixed(2)}</td>
          <td>${huruf}</td>
        `;
    hasilBody.appendChild(hasilRow);
  });
}
function konversiHuruf(rata) {
  if (rata >= 85) return "A";
  if (rata >= 70) return "B";
  if (rata >= 60) return "C";
  if (rata >= 50) return "D";
  return "E";
}
function exportKeExcel() {
  const rows = document.querySelectorAll("#tabelMahasiswa tbody tr");
  if (rows.length === 0) return alert("Belum ada data untuk diekspor.");
  const header = ["Nama"];
  for (let i = 1; i <= jumlahPertemuan; i++) header.push(`P${i}`);
  header.push("Total", "Rata-rata", "Huruf");
  const data = [header];
  rows.forEach((row) => {
    const nama = row.querySelector(".nama").value.trim();
    const nilaiInputs = row.querySelectorAll(".nilai");
    let nilai = [];
    let total = 0;
    let valid = true;
    nilaiInputs.forEach((input) => {
      const val = parseFloat(input.value);
      if (isNaN(val)) valid = false;
      else {
        nilai.push(val);
        total += val;
      }
    });
    if (!valid || nama === "") return;
    const rata = total / jumlahPertemuan;
    const huruf = konversiHuruf(rata);
    data.push([nama, ...nilai, total.toFixed(2), rata.toFixed(2), huruf]);
  });
  const worksheet = XLSX.utils.aoa_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Nilai Mahasiswa");
  XLSX.writeFile(workbook, "nilai_mahasiswa_lengkap.xlsx");
}
function importDariExcel(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    if (json.length <= 1) return;
    jumlahPertemuan = json[0].length - 4;
    const tbody = document.querySelector("#tabelMahasiswa tbody");
    tbody.innerHTML = "";
    const theadRow = document.querySelector("#tabelMahasiswa thead tr");
    theadRow.innerHTML = "<th>Nama</th>";
    for (let i = 1; i <= jumlahPertemuan; i++) {
      const th = document.createElement("th");
      th.textContent = `P${i}`;
      theadRow.appendChild(th);
    }
    theadRow.innerHTML += "<th>Aksi</th>";
    for (let i = 1; i < json.length; i++) {
      const row = json[i];
      if (!row[0]) continue;
      const tr = document.createElement("tr");
      const tdNama = document.createElement("td");
      const inputNama = document.createElement("input");
      inputNama.type = "text";
      inputNama.className = "nama";
      inputNama.value = row[0];
      tdNama.appendChild(inputNama);
      tr.appendChild(tdNama);
      for (let j = 1; j <= jumlahPertemuan; j++) {
        const tdNilai = document.createElement("td");
        const inputNilai = document.createElement("input");
        inputNilai.type = "number";
        inputNilai.className = "nilai";
        inputNilai.value = row[j] ?? "";
        tdNilai.appendChild(inputNilai);
        tr.appendChild(tdNilai);
      }
      const tdAksi = document.createElement("td");
      const btnHapus = document.createElement("button");
      btnHapus.textContent = "Hapus";
      btnHapus.className = "btn-hapus";
      btnHapus.onclick = () => tr.remove();
      tdAksi.appendChild(btnHapus);
      tr.appendChild(tdAksi);
      tbody.appendChild(tr);
    }
    aktifkanAutoTab();
  };
  reader.readAsArrayBuffer(file);
}
function aktifkanAutoTab() {
  document
    .querySelector("#tabelMahasiswa")
    .addEventListener("keydown", function (e) {
      if (e.target.classList.contains("nilai") && e.key === "Enter") {
        e.preventDefault();
        const input = e.target;
        const td = input.parentElement;
        const tr = td.parentElement;
        const semuaInput = Array.from(tr.querySelectorAll("input.nilai"));
        const indexSekarang = semuaInput.indexOf(input);
        const inputSelanjutnya = semuaInput[indexSekarang + 1];
        if (inputSelanjutnya) inputSelanjutnya.focus();
      }
    });
}
document.addEventListener("DOMContentLoaded", aktifkanAutoTab);
  </script>
</body>
</html>
