<template>
  <div>
    <h2>导出Word</h2>
    <div class="content">
      <line-chart :xData="barChartData.xData" :chartData="barChartData.data"></line-chart>
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
import PizZipUtils from 'pizzip/utils/index.js'
import { saveAs } from 'file-saver'

function loadFile (url, callback) {
  PizZipUtils.getBinaryContent(url, callback)
}

function replaceErrors (key, value) {
  if (value instanceof Error) {
    return Object.getOwnPropertyNames(value).reduce(function (
      error,
      key
    ) {
      error[key] = value[key]
      return error
    },
    {})
  }
  return value
}

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
      ]
    }
  },
  methods: {
    renderDoc () {
      const ImageModule = require('docxtemplater-image-module-free')
      loadFile('docxTemplate/template.docx', function (
        error,
        content
      ) {
        if (error) {
          throw error
        }
        /**
         * 添加图片Module
         */
        var opts = {}
        opts.centered = false
        opts.getImage = function (tagValue, tagName) {
          return new Promise(function (resolve, reject) {
            PizZipUtils.getBinaryContent(tagValue, function (error, content) {
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
        var doc = new Docxtemplater(zip, { modules: [imageModule] })
        doc.setData({
          first_name: 'John',
          last_name: 'Doe',
          phone: '0652455478',
          description: 'New Website'
        })
        try {
          // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
          doc.render()
        } catch (error) {
          // The error thrown here contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).

          console.log(JSON.stringify({ error: error }, replaceErrors))
          if (error.properties && error.properties.errors instanceof Array) {
            const errorMessages = error.properties.errors
              .map(function (error) {
                return error.properties.explanation
              })
              .join('\n')
            console.log('errorMessages', errorMessages)
            // errorMessages is a humanly readable message looking like this :
            // 'The tag beginning with "foobar" is unopened'
          }
          throw error
        }
        var out = doc.getZip().generate({
          type: 'blob',
          mimeType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        }) // Output the document using Data-URI
        saveAs(out, 'output.docx')
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
