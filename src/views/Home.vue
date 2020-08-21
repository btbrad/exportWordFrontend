<template>
  <div class="home">
    <export-word />
  </div>
</template>

<script>
// @ is an alias to /src
import ExportWord from '@/components/exportWord/index'

export default {
  name: 'Home',
  components: {
    ExportWord
  }
}
</script>
