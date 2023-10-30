<template> 
  <h1 class="title is-3" >Convert spreadsheet cells to Markdown</h1>
  <h2 class="subtitle">Copy & Paste your Excel Cells into the textfield below. It will be converterd to a Markdown Table.</h2>
  <MarkdownConverterComponent />
</template>

<script>
import MarkdownConverterComponent from './components/MarkdownConverterComponent.vue';

export default {
  name: 'App',
  components: {
    MarkdownConverterComponent
  }
}
</script>

<style>
@import "~bulma/css/bulma.css";
</style>
