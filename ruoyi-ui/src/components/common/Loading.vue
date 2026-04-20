<template>
    <div class="loading-mask" v-if="visible">
        <div class="loading-content">
            <el-loading-spinner class="loading-spinner"></el-loading-spinner>
            <p class="loading-text">{{ text }}</p>
        </div>
    </div>
</template>

<script>
export default {
  name: 'Loading',
  props: {
    visible: {
      type: Boolean,
      default: false
    },
    text: {
      type: String,
      default: '???...'
    }
  }
}
</script>

<style scoped>
    .loading-mask {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(255, 255, 255, 0.7);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 9999;
    }

    .loading-content {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 10px;
    }

    .loading-text {
        color: #409eff;
        font-size: 14px;
    }
</style>