<template>
  <div class="submenu">
    <el-menu
      default-active="1"
      class="el-menu-vertical-demo"
      active-text-color="#ffd04b"
      background-color="#545c64"
      text-color="#fff"
      @open="handleOpen"
      @close="handleClose"
    >
      <el-menu-item index="1">
        <el-icon><i-ep-circle-check-filled /></el-icon>
        <span>模板生成</span>
      </el-menu-item>
    </el-menu>
  </div>
</template>

<script setup></script>

<style lang="scss" scoped>
.el-menu {
  border-right: 0;
}
.submenu {
  height: 100%;
  border-right: 1px solid #dddada;
}
</style>
