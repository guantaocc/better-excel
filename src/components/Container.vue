<template>
  <div class="container">
    <SideBar class="sidebar-container"></SideBar>
    <Main class="main-container"></Main>
  </div>
</template>

<script setup>
import SideBar from './SideBar.vue'
import Main from './Main.vue'
</script>

<style lang="scss" scoped>
.container {
  width: 100%;
  height: 100%;
  box-sizing: border-box;
  display: flex;
  .sidebar-container {
    width: 256px;
  }
  .main-container {
    flex: 1;
    box-sizing: border-box;
    padding:1rem;
    height: 100%;
    overflow-y: auto;
  }
}
</style>