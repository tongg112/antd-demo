<template>
  <div id="app">
    <viewer v-viewer="{toolbar: false}">
      <img alt="Vue logo" src="./assets/logo.png">
    </viewer>
<!--    <HelloWorld msg="Welcome to Your Vue.js App"/>-->
    <LuckysheetDemo msg="Welcome to Your Vue.js App"/>
  </div>
</template>

<script>
import LuckysheetDemo from './components/LuckysheetDemo.vue'

export default {
  name: 'App',
  components: {
    LuckysheetDemo
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
