<template>
  <div>
    <h2>导出Word</h2>
    <div class="content">
      <line-chart ref="chart" :xData="barChartData.xData" :chartData="barChartData.data"></line-chart>
      <el-table :data="tableData" border>
        <el-table-column prop="name" label="姓名" align="center"></el-table-column>
        <el-table-column prop="age" label="年龄" align="center"></el-table-column>
        <el-table-column prop="gender" label="性别" align="center"></el-table-column>
      </el-table>
    </div>
    <el-button type="primary" icon="el-icon-document" @click="renderDoc">导出Word</el-button>
  </div>
</template>

<script>
import LineChart from '@/components/charts/lineCharts'
import Docxtemplater from 'docxtemplater'
import PizZip from 'pizzip'
import JSZipUtils from 'jszip-utils'
import { saveAs } from 'file-saver'

const ImageModule = require('docxtemplater-image-module-free')

export default {
  name: 'ExportWord',
  components: {
    LineChart
  },
  data () {
    return {
      barChartData: {
        xData: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
        data: [820, 932, 901, 934, 1290, 1330, 1320]
      },
      tableData: [
        { name: '小明', age: '23', gender: '男' },
        { name: '小红', age: '17', gender: '女' },
        { name: '小花', age: '28', gender: '女' }
      ],
      image: ''
    }
  },
  methods: {
    getImage () {
      this.image = this.$refs.chart.image
    },
    renderDoc () {
      this.getImage()
      const _this = this
      JSZipUtils.getBinaryContent('docxTemplate/template.docx', function (error, content) {
        if (error) {
          console.error(error)
          return
        }
        var opts = {}
        opts.centered = false
        opts.getImage = function (tagValue, tagName) {
          return new Promise(function (resolve, reject) {
            JSZipUtils.getBinaryContent(tagValue, function (error, content) {
              if (error) {
                return reject(error)
              }
              return resolve(content)
            })
          })
        }
        opts.getSize = function (img, tagValue, tagName) {
        // FOR FIXED SIZE IMAGE :
          // return [150, 150]

          // FOR IMAGE COMING FROM A URL (IF TAGVALUE IS AN ADRESS) :
          // To use this feature, you have to be using docxtemplater async
          // (if you are calling setData(), you are not using async).
          return new Promise(function (resolve, reject) {
            var image = new Image()
            image.src = tagValue
            image.onload = function () {
              resolve([image.width, image.height])
            }
            image.onerror = function (e) {
              console.log('img, tagValue, tagName : ', img, tagValue, tagName)
              alert('An error occured while loading ' + tagValue)
              reject(e)
            }
          })
        }

        var imageModule = new ImageModule(opts)

        var zip = new PizZip(content)
        var doc = new Docxtemplater()
          .loadZip(zip)
          .attachModule(imageModule)
          .compile()

        doc.resolveData({
          image: _this.image,
          students: _this.tableData
        }).then(function () {
          console.log('ready')
          doc.render()
          var out = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          })
          saveAs(out, 'generated.docx')
        })
      })
    }
  }
}
</script>

<style scoped lang="scss">
  .content {
    width: 1000px;
    height: auto;
    margin: 50px auto;
    border: 1px solid #ccc;
  }
</style>
