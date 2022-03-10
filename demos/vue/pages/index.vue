<!-- xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com -->
<template>
<div>
  <div v-for="item in data.data" v-bind:key="item.name">
    <h2>{{ item.name }}</h2>
    <table>
      <tr v-for="row in item.data" v-bind:key="row.Index">
        <td>{{ row.Name }}</td>
        <td>{{ row.Index }}</td>
      </tr>
    </table>
  </div>
</div>
</template>

<script>
export default {
  async asyncData ({$content}) {
    return {
      data: await $content('sheetjs').fetch()
    };
  }
}
</script>
