<template>
  <input type="file" name="excel" id="excel" @change="loadExcel" />
  <select v-if="spreadsheet != null" @change="changeSheet">
    <template v-for="(sheet, ix) in spreadsheet.workbook.worksheets">
      <option :value="ix">{{ sheet.name }}</option>
    </template>
  </select>
  <hr />
  <table v-if="activeWorksheet != null">
    <tr v-for="(_, i) in activeWorksheet.dimensions.to.row + 1">
      <td v-for="(__, j) in activeWorksheet.dimensions.to.col + 1">
        {{ activeWorksheet.cells[j][i] }}
      </td>
    </tr>
    {{
      `Total: ${activeWorksheet.dimensions.to.col}:${activeWorksheet.dimensions.to.row}`
    }}
  </table>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import { Spreadsheet } from './lib/Spreadsheet';
import { Worksheet } from './lib/Worksheet';

let spreadsheet = ref<Spreadsheet | undefined>(undefined);
let activeWorksheet = ref<Worksheet | undefined>(undefined);

function changeSheet(e: Event) {
  activeWorksheet.value = spreadsheet.value!.workbook.worksheets[parseInt((e.target! as HTMLSelectElement).value)];
}

async function loadExcel(e: Event) {
  const target = e.target as HTMLInputElement;
  if (target.files == null || target.files.length <= 0) return;

  spreadsheet.value = await Spreadsheet.loadFromFile(target.files[0]);
  activeWorksheet.value = spreadsheet.value.workbook.worksheets[0];
}
</script>

<style scoped>
.numeric {
  text-align: right;
}

table {
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  font-size: small;
  width: 100%;
  border-collapse: collapse;
  border: 2px solid black;
}

th,
td {
  padding: 2px 4px;
  border: 1px solid black;
}
</style>
