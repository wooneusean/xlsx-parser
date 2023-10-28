<template>
  <input type="file" name="excel" id="excel" @change="loadExcel" />
  <h4 v-if="spreadsheet != null">This spreadsheet has {{ spreadsheet.workbook?.worksheets.length }} worksheets.</h4>
</template>

<script setup lang="ts">
import { ref } from 'vue';
import Spreadsheet from './components/classes/Spreadsheet';

let spreadsheet = ref<Spreadsheet | undefined>(undefined);

async function loadExcel(e: Event) {
  const target = e.target as HTMLInputElement;
  if (target.files == null || target.files.length <= 0) return;

  spreadsheet.value = await Spreadsheet.loadFromFile(target.files[0]);
  console.log(spreadsheet.value);
}
</script>

<style scoped></style>
