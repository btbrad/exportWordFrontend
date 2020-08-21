<template>
  <div class="line-charts" style="width: 1000px;height: 500px;"></div>
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
      chart: null,
      image: null
    }
  },
  mounted () {
    this.init()
    this.$nextTick(() => {
      this.getImage()
    })
  },
  methods: {
    init () {
      this.chart = echarts.init(this.$el)
      this.setOptions()
    },
    setOptions () {
      this.chart.setOption({
        animation: false,
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
    },
    getImage () {
      this.image = this.chart.getDataURL({
        type: 'png',
        pixelRatio: 2,
        backgroundColor: '#fff'
      })
      // this.image = this.$el.querySelector('canvas').toDataURL()
    }
  }
}
</script>

<style scoped lang="scss">

</style>
