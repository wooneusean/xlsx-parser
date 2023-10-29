<template>
  <input type="file" name="excel" id="excel" @change="loadExcel" />
  <h4 v-if="spreadsheet != null">This spreadsheet has {{ spreadsheet.workbook.worksheets.length }} worksheets.</h4>

  <input type="text" name="cell-reference" id="cell-reference" @change="getValueAtReference" placeholder="reference, e.g A4, P701" />
  Value at reference is: {{ referenceValue }}

  <ul v-if="spreadsheet != null">
    <li v-for="worksheet in spreadsheet.workbook.worksheets">
      {{ worksheet.name }}
      <ul>
        <li v-for="row in worksheet.rows">
          Row #{{ row.index }}
          <ul>
            <li v-for="cell in row.cells">{{ cell.reference }}: {{ cell.value }}</li>
          </ul>
        </li>
      </ul>
    </li>
  </ul>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import Spreadsheet from './components/classes/Spreadsheet';

let spreadsheet = ref<Spreadsheet | undefined>(undefined);
let referenceValue = ref<string>('');

function getValueAtReference(e: Event) {
  referenceValue.value =
    spreadsheet.value?.workbook.worksheets[0].at((e.target as HTMLInputElement).value)?.value ?? 'null';
}

async function loadExcel(e: Event) {
  const target = e.target as HTMLInputElement;
  if (target.files == null || target.files.length <= 0) return;

  spreadsheet.value = await Spreadsheet.loadFromFile(target.files[0]);
}
</script>

<style scoped></style>
