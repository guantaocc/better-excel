<template>
  <div class="main">
    <el-card>
      <template #header>
        <div class="card-header">
          <span>文件选择</span>
        </div>
      </template>
      <el-button type="primary" @click="openFile">请选择文件(支持xls和xlsx文件)</el-button>
      <div>{{selectPath}}</div>
    </el-card>

    <el-card class="template">
      <template #header>
        <div class="card-header">
          <span>模板粘贴</span>
           <el-alert title="请替换模板 {{ }} 中的字符为目标表格中 每一列的表头名称" type="info" />
        </div>
      </template>
      <el-input
        v-model="textarea"
        style="width: 100%"
        :autosize="{ minRows: 6, maxRows: 10 }"
        type="textarea"
        placeholder="示例：UtranCell={{Utrancell}}"
      />
    </el-card>
    <div class="gen">
      <span>请输入第几个工作表：<el-input style="width:100px" v-model="sheetIndex"></el-input></span>
      <el-button style="margin-left:10px" type="primary" @click="gen">生成指令</el-button>
      <el-button style="margin-left:10px" type="primary" @click="download">下载</el-button>
    </div>
  </div>
</template>

<script setup>

const selectPath = ref('')
const textarea = ref('')
const downloadMessage = ref('')
const sheetIndex = ref(2)

const openFile = async () => {
  const res = await ipcRenderer.invoke('openFileDialog')
  if(res){
    selectPath.value = res
  }
}

const gen = async () => {
  if(!textarea.value){
    ElMessage({
    message: '模板不能为空!',
    type: 'warning',
  })
    return
  }
  if(!selectPath.value){
    ElMessage('目标文件不能为空!')
    return
  }
  const { errorMessage, result } = await ipcRenderer.invoke('genTplFromExcel', selectPath.value, textarea.value, sheetIndex.value)
  console.log(errorMessage, result);
  if(!errorMessage){
    ElMessage({
      message: '生成成功，请下载记录',
      type: 'success'
    })
    downloadMessage.value = result
  }else {
    ElMessage({
      message: errorMessage,
      type: 'error'
    })
  }
}

const download = async () => {
  if(!downloadMessage.value){
    ElMessage({
      message: '请先生成指令',
      type: 'warning'
    })
    return
  }
  // 发送主进程下载
  const {success, filePath} = await ipcRenderer.invoke('downloadFileDialog', downloadMessage.value)
  if(success){
      ElMessageBox.alert(`文件地址：${filePath}`, '下载成功', {
      // if you want to disable its autofocus
      // autofocus: false,
      confirmButtonText: 'OK',
    })
  }else {
    ElMessage({
      message: '下载异常',
      type: 'error'
    })
  }
}

</script>

<style lang="scss" scoped>
.template {
  margin-top: 1rem;
}
.gen {
  margin-top: 20px;
}
</style>