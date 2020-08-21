<template>
  <div class="home">
    <export-word />
    <line-chart :xData="barChartData.xData" :chartData="barChartData.data"></line-chart>
  </div>
</template>

<script>
// @ is an alias to /src
import ExportWord from '@/components/exportWord/index'
import LineChart from '@/components/charts/lineCharts'

export default {
  name: 'Home',
  components: {
    ExportWord,
    LineChart
  },
  data () {
    return {
      barChartData: {
        xData: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
        data: [820, 932, 901, 934, 1290, 1330, 1320]
      }
    }
  }
}
</script>
