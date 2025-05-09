<template>
  <div class="container">
    <h1>Extracteur PDF vers Excel pour offre Legrand</h1>
    <div class="upload-section">
      <label for="client">Numéro de client : </label>
      <input type="text" name="client" id="client" v-model="client">
    </div>
    <div class="upload-section">
      <label for="user">Vos initiales : </label>
      <input type="text" name="user" id="user" placeholder="ex: nse" v-model="user">
    </div>
    <div class="upload-section" v-if="client && user">
      <input type="file" accept=".pdf" @change="handleFileUpload"/>
    </div>

    <div v-if="extractedData.length" class="results-section">
      <h2>Données extraites</h2>
      <table>
        <thead>
        <tr>
          <th>Référence</th>
          <th>Quantité</th>
        </tr>
        </thead>
        <tbody>
        <tr v-for="(item, index) in extractedData" :key="index">
          <td>{{ item.reference }}</td>
          <td>{{ item.quantity }}</td>
        </tr>
        </tbody>
      </table>

      <p>Référence trouvé dans le PDF : {{ pdfReferenceFound }}</p>
      <p>Référence matché avec référence Legrand : {{ matchedReferenceFound }}</p>
      <div v-if="itemNotFound.length > 0">
        <p>Référence non trouvé : </p>
        <ul>
          <li v-for="item in itemNotFound">{{ item }}</li>
        </ul>
      </div>
      <button @click="exportToExcel" class="export-btn">
        Exporter vers Excel
      </button>
    </div>
  </div>
</template>

<script setup lang="ts">
import {ref} from 'vue'
import * as XLSX from "xlsx"
import * as pdfjsLib from 'pdfjs-dist'

pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`

interface ExtractedItem {
  client: string
  reference: string
  quantity: number
}

const pdfReferenceFound = ref(0);
const matchedReferenceFound = ref(0);
const itemNotFound = ref([]);
const user = ref();
const client = ref();

const extractedData = ref<ExtractedItem[]>([])

async function handleFileUpload(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (!file) return;

  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;

  const extractedItems: ExtractedItem[] = [];

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const text = textContent.items.map((item: any) => item.str).join(" ");

    const cleanedText = text.replace(/\s+/g, ' ').trim();
    const regex = /([A-Z0-9]{3,20})\s+(\d{1,3}(?:\s?\d{3})?)+(,00)\s+(pc|m|Piece|Meter|piece|meter)/g;
    const data = await getData();
    let match;
    while ((match = regex.exec(cleanedText)) !== null) {
      let reference = match[1];
      const quantity = match[2].replace(/\s+/g, '')
      pdfReferenceFound.value++;
      if (!isNaN(quantity) && quantity > 0) {
        const matchItems = data.find(item => {
          if (item["Ref.Legrand"].toString() === reference) {
            reference = item["Ref.Rexel"];
            matchedReferenceFound.value++;
            return true;
          } else if (reference.startsWith("VG")) {
            if (item["Ref.Legrand"].toString() === reference.slice(2)) {
              reference = item["Ref.Rexel"];
              matchedReferenceFound.value++;
              return true;
            }
          }
          return false;
        });

        if (!matchItems) {
          itemNotFound.value.push(reference);
        }

        extractedItems.push({client: client.value, reference, quantity});
      }
    }
  }
  extractedData.value = extractedItems;
}

async function getData() {
  try {
    const filePath = "/data/REF-LEGRAND.xlsx";
    const response = await fetch(filePath);
    if (!response.ok) {
      throw new Error(`Impossible de charger le fichier : ${response.statusText}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(new Uint8Array(arrayBuffer), {type: "array"});
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json<{ "Ref_Rexel": string; "Ref_Legrand": string }>(worksheet);
  } catch (error) {
    console.error("Erreur lors de la lecture du fichier Excel :", error);
  }
}

function exportToExcel() {
  const worksheetData = [
    ["char(11)", "char(35)", "num(15,3)"],
    ["cunum", "cuprdc", "cuqty"],
    ...extractedData.value.map(item => [item.client, item.reference, item.quantity])
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, `reoias${user.value}`);
  XLSX.writeFile(workbook, `reoias${user.value}.csv`);
}
</script>

<style>
.container {
  max-width: 800px;
  margin: 0 auto;
  padding: 2rem;
}

.upload-section {
  margin: 2rem 0;
}

.results-section {
  margin-top: 2rem;
}

table {
  width: 100%;
  border-collapse: collapse;
  margin: 1rem 0;
}

th, td {
  border: 1px solid #ddd;
  padding: 0.5rem;
  text-align: left;
}

th {
  background-color: #f5f5f5;
}

.export-btn {
  background-color: #4CAF50;
  color: white;
  padding: 0.5rem 1rem;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.export-btn:hover {
  background-color: #45a049;
}
</style>
