<template>
  <div class="demo">
    <div class="header">
      <!-- 本地导入 -->

      <a-upload :file-list="fileList" :before-upload="beforeUpload" accept=".xlsx,.xls">
        <a-button> <a-icon type="upload" /> 点击导入 </a-button>
      </a-upload>
      <a-button style="margin-left: 20px" @click="getData">请求数据</a-button>
      <a-button style="margin-left: 20px" @click="downloadExcel">点击导出</a-button>
    </div>

    <!-- luckysheet容器 -->
    <div
        id="luckysheet"
        style="margin: 0px; padding: 0px; position: absolute; width: 96%; left: 0px; top: 400px; bottom: 0px; z-index: 0"
    ></div>
  </div>
</template>

<script>
import LuckyExcel from 'luckyexcel'
import { exportExcel } from '@/utils/export'

export default {
  name: 'LuckysheetDemo',
  data () {
    return {
      fileList: [],
      name: '导出的表格',
      options: {
        container: 'luckysheet',
        title: '测试Excel', // 表 头名
        lang: 'zh', // 中文
        showtoolbar: true, // 是否显示工具栏
        showinfobar: false, // 是否显示顶部信息栏
        showsheetbar: true // 是否显示底部sheet按钮
      }
    }
  },
  mounted () {
    window.luckysheet.create(this.options)
  },
  methods: {
    // 导出到本地
    downloadExcel () {
      exportExcel(window.luckysheet.getAllSheets(), this.name)
    },
    // 导入数据
    beforeUpload (files) {
      console.log(files)
      if (files === null || files.length === 0) {
        // alert('No files wait for import')
        this.$message.error('文件为空', 2.5)
        return
      }
      const name = files.name
      const suffixArr = name.split('.')
      this.name = suffixArr[0]
      const suffix = suffixArr[suffixArr.length - 1]
      if (suffix !== 'xlsx') {
        this.$message.error('请传入后缀为 xlsx 的文件', 2.5)
        return
      }
      LuckyExcel.transformExcelToLucky(files, (exportJson) => {
        if (exportJson.sheets === null || exportJson.sheets.length === 0) {
          this.$message.error('无法读取excel文件的内容，当前不支持xls文件！')
          return
        }
        window.luckysheet.destroy()
        this.options.data = exportJson.sheets
        console.log('数据', this.options.data)
        window.luckysheet.create(this.options)
      })
    },
    getData () {
      const data = [
        {
          name: 'cell',
          color: '',
          index: 1,
          status: 0,
          order: 1,
          celldata: [
            {
              r: 0, // 行
              c: 0, // 列
              v: '姓名' // 值
            },
            {
              r: 1, // 行
              c: 0, // 列
              v: '张三' // 值
            },
            {
              r: 2, // 行
              c: 0, // 列
              v: '李四' // 值
            },
            {
              r: 3, // 行
              c: 0, // 列
              v: '王五' // 值
            },
            {
              r: 0,
              c: 1,
              v: '年龄'
            },
            {
              r: 1,
              c: 1,
              v: '1'
            },
            {
              r: 2,
              c: 1,
              v: '2'
            },
            {
              r: 3,
              c: 1,
              v: '3'
            }
          ],
          config: {}
        },
        {
          name: 'Sheet2',
          color: '',
          index: 1,
          status: 0,
          order: 1,
          celldata: [],
          config: {}
        },
        {
          name: 'Sheet3',
          color: '',
          index: 2,
          status: 0,
          order: 2,
          celldata: [],
          config: {}
        }
      ]
      window.luckysheet.destroy()
      this.options.data = data
      window.luckysheet.create(this.options)
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style>
* {
  margin: 0;
  padding: 0;
}
.header {
  display: flex;
  margin-bottom: 20px;
}
.printHideCss {
  visibility: hidden !important;
}
.demo {
  width: 100%;
}
</style>
