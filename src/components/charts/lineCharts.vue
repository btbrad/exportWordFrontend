<template>
  <div class="line-charts" style="width: 500px;height: 500px;"></div>
</template>

<script>
import echarts from 'echarts'

export default {
  name: 'LineCharts',
  props: {
    xData: {
      type: Array,
      default: () => []
    },
    chartData: {
      type: Array,
      default: () => []
    }
  },
  data () {
    return {
      chart: null
    }
  },
  mounted () {
    this.init()
  },
  methods: {
    init () {
      this.chart = echarts.init(this.$el)
      this.setOptions()
    },
    setOptions () {
      this.chart.setOption({
        xAxis: {
          type: 'category',
          data: this.xData
        },
        yAxis: {
          type: 'value'
        },
        series: [{
          data: this.chartData,
          type: 'line'
        }]
      })
    }
  }
}
</script>

<style scoped lang="scss">

</style>
