<script setup>
import { ref, reactive, onMounted, onUnmounted } from 'vue'
import { downloadFile, exportReport} from '@/api/index'
import { buildExportBody} from '@/utils/request'
import { usePolling} from '@/composables/usePolling'

const {
  startPolling,
  stopPolling
} = usePolling()
/* ───── 认证 ───── */
const token = ref('')

/* ───── 按钮配置 ───── */
const EXPORT_BUTTONS = [
  { type: 'total',                      label: '常规订单数据' },
  { type: 'fbk45Section',               label: '方案制单耗时数据' },
  { type: 'programmeApproveTimeSection',label: '方案审批耗时数据' },
  { type: 'enquiryApproveTimeSection',  label: '比价单审批耗时数据' },
  { type: 'orderApproveTimeSection',    label: '订单审批耗时数据' },
  { type: 'bidApproveTimeSection',      label: '定标流程耗时数据' },
  { type: 'orderFrameInfo',             label: '框架合同下单数据' }
];

/* ───── 每个导出任务的独立状态 ───── */
// exporting: 导出请求进行中;
const taskState = reactive(
  Object.fromEntries(EXPORT_BUTTONS.map(b => [b.type, {
    taskId: null,
    // exporting: false, 
    state: '',   // generating | generated | downloading | processing | success | error
    downloadUrl: '',    // 文件下载链接
    timerId: null,      // setInterval id
    alertMsg: '',
    alertType: 'info',
  }]))
);

/* ───── 全局错误（token 校验等） ───── */
const alertMsg = ref('');
const alertType = ref('info');
const errors = ref({});

/* ───── 校验 ───── */
function validate() {
  const e = {};
  if (!token.value.trim()) e.token = '请输入 token';
  errors.value = e;
  return Object.keys(e).length === 0;
}

/* ───── 设置任务级提示 ───── */
function setTaskAlert(type, alertType, msg) {
  taskState[type].alertType = alertType;
  taskState[type].alertMsg = msg;
}

/* ───── 导出数据 ───── */
async function fetchData(type = 'total') {
  if (!validate()) {
    alertType.value = 'error';
    alertMsg.value = '请先填写认证 Token';
    return;
  }

  if(type === 'orderFrameInfo') {
    exportOrderFrameInfo()
    return 
  }

  alertMsg.value = '';
  const state = taskState[type];

  // 若已在轮询中，不重复发起
  if (state.state === 'generating') return;
  state.taskId = null;
  // state.exporting = true;
  state.state = 'generating';
  state.alertMsg = '';
  state.downloadUrl = '';

  try {
    const res = await exportReport(
      token.value,
      buildExportBody(type)
    )

    console.log('[export] resp', res);
    if(res.data.status !== '0') {
      // state.exporting = false;
      state.state = 'error';
      return Promise.reject(new Error(res.data.message || '导出请求失败'))
    }
      

    // 导出请求成功，启动轮询
    setTaskAlert(type, 'info', '导出任务已提交，正在等待服务器生成文件…');
    setTimeout(() => {
      startPolling(
        type,
        token.value,
        (res) => {
          const data = res.data.blocks.resultBlock.data
          
          const taskId = state.taskId
          if(!taskId) {
            const task = data.find(item => item.fbk4 == '1')
            if (task) {
              state.taskId = task.id
              
            } else {
              // state.exporting = false
              state.state = 'error'
              setTaskAlert(type, 'error', '未找到对应的导出任务，请确认业务系统已正确处理请求')
            }
          }
          else {
            const task = data.find(item => item.id == taskId)
            if (task) {
              if (task.fbk4 == '2') {
                state.state = 'generated'
                // state.exporting = false
                state.downloadUrl = task.attachmentLink
                console.log(`[download] ${type}`, state.downloadUrl);
                setTaskAlert(type, 'success', '文件已生成，点击下载按钮获取文件')
                stopPolling(type)
              } else {
                console.log(`任务 ${taskId} 状态：${task.fbk4}，继续轮询...`)
              }
            }

          }
        },
        (err) => {
          console.error(err)
        }
      )
    }, 1000)
    
  } catch (err) {
    console.error(err);
    let msg = `请求失败：${err.message}`;
    setTaskAlert(type, 'error', msg);
    // state.exporting = false;
    state.state = 'error';
  }
}

const exportOrderFrameInfo = async () => {
  const res = await downloadFile({
    token: token.value,
    url: 'http://172.22.35.101:8001/zjark-cloud-srmreport/orderFrameInfo/exportExcel',
    fileName: `框架合同下单数据-${new Date().getMonth() + 1}月.xlsx`,
    taskType: 'orderFrameInfo',
    method: 'POST',
    data: buildExportBody('orderFrameInfo'),
  })

  if (!res.success) {
    state.state = 'error'

    setTaskAlert(
      type,
      'error',
      `下载失败：${res.message}`
    )
  }
}

const download = async (type, label) => {
  const url = taskState[type].downloadUrl

  if (!url) return

  const state = taskState[type]

  state.state = 'downloading'

  setTaskAlert(
    type,
    'info',
    '正在下载文件...'
  )

  const res = await downloadFile({
    url,
    fileName: `${label}-${new Date().getMonth() + 1}月.xlsx`,
    taskType: type
  })

  if (!res.success) {
    state.state = 'error'

    setTaskAlert(
      type,
      'error',
      `下载失败：${res.message}`
    )
  }
}

onMounted(() => {
  window.electronAPI.onDownloadStatus((data) => {
    const state = taskState[data.type]

    if (!state) return

    switch (data.stage) {

      case 'processing':
        state.state = 'processing'
        setTaskAlert(
          data.type,
          'info',
          '文件下载完成，正在处理Excel数据...'
        )
        break

      case 'success':
        state.state = 'success'
        setTaskAlert(
          data.type,
          'success',
          'Excel处理完成，已保存到本地'
        )
        break

      case 'error':
        state.state = 'error'
        setTaskAlert(
          data.type,
          'error',
          data.message || '处理失败'
        )
        break
    }
  })
})

onUnmounted(() => {
  window.electronAPI.removeDownloadStatus()
})

</script>

<template>
  <!-- 顶部导航 -->
  <div class="header">
    <span class="header-icon">📦</span>
    <div>
      <h1>采购跟踪报表工具</h1>
      <p>输入 Token → 配置查询条件 → 获取报表数据</p>
    </div>
  </div>

  <div class="container">

    <!-- ① Token 配置 -->
    <div class="card">
      <div class="card-header">
        <span class="icon">🔑</span>
        身份认证 Token
      </div>
      <div class="card-body">
        <div class="token-grid">
          <div>
            <div class="field-label">
              token <span class="tag">Header</span>
            </div>
            <input
              v-model="token"
              type="text"
              placeholder="粘贴 token Header 值"
              :class="{ error: errors.token }"
            />
            <div v-if="errors.token" style="color:#ff4d4f;font-size:12px;margin-top:4px">{{ errors.token }}</div>
          </div>
        </div>
      </div>
    </div>

    <!-- ③ 操作按钮 -->
    <div class="card">
      <div class="card-header">
        <span class="icon">🚀</span>
        导出任务
      </div>
      <div class="card-body">

        <!-- 全局 token 错误提示 -->
        <div v-if="alertMsg" :class="['alert', 'alert-' + alertType]" style="margin-bottom:16px">
          <span>{{ alertType === 'error' ? '❌' : 'ℹ️' }}</span>
          <span>{{ alertMsg }}</span>
        </div>

        <!-- 两列布局：左导出 / 右下载 -->
        <div class="export-grid">

          <!-- 每行：导出按钮 + 对应下载按钮 + 行内提示 -->
          <template v-for="btn in EXPORT_BUTTONS" :key="btn.type">

          <!-- 框架合同下单数据：只显示一个按钮 -->
          <template v-if="btn.type === 'orderFrameInfo'">

            <div class="export-cell export-cell-span">
              <button
                class="btn btn-primary btn-full"
                @click="fetchData(btn.type)"
              >
                ⬇️ {{ btn.label }}
              </button>

              <div
                v-if="taskState[btn.type].alertMsg"
                :class="['alert', 'alert-' + taskState[btn.type].alertType]"
                style="margin-top:6px;padding:6px 10px;font-size:12px"
              >
                <span>
                  {{
                    taskState[btn.type].alertType === 'error'
                      ? '❌'
                      : taskState[btn.type].alertType === 'success'
                      ? '✅'
                      : 'ℹ️'
                  }}
                </span>
                <span>{{ taskState[btn.type].alertMsg }}</span>
              </div>
            </div>

          </template>

          <!-- 其它报表：保持导出+下载两列 -->
          <template v-else>

            <!-- 导出按钮 -->
            <div class="export-cell">
              <button
                class="btn btn-primary btn-full"
                :disabled="taskState[btn.type].state === 'generating'"
                @click="fetchData(btn.type)"
              >
                <span
                  v-if="taskState[btn.type].state === 'generating'"
                  class="spinner"
                ></span>
                <span v-else>🚀</span>

                <span v-if="taskState[btn.type].state === 'generating'">
                  等待生成…
                </span>
                <span v-else>
                  {{ btn.label }}
                </span>
              </button>

              <div
                v-if="taskState[btn.type].alertMsg"
                :class="['alert', 'alert-' + taskState[btn.type].alertType]"
                style="margin-top:6px;padding:6px 10px;font-size:12px"
              >
                <span>
                  {{
                    taskState[btn.type].alertType === 'error'
                      ? '❌'
                      : taskState[btn.type].alertType === 'success'
                      ? '✅'
                      : 'ℹ️'
                  }}
                </span>
                <span>{{ taskState[btn.type].alertMsg }}</span>
              </div>
            </div>

            <!-- 下载按钮 -->
            <div class="export-cell">
              <button
                class="btn btn-success btn-full"
                :disabled="!taskState[btn.type].downloadUrl"
                @click="download(btn.type, btn.label)"
              >
                <span
                  v-if="
                    taskState[btn.type].state === 'downloading' ||
                    taskState[btn.type].state === 'processing'
                  "
                  class="spinner"
                ></span>
                <span v-else>⬇️</span>

                {{
                  taskState[btn.type].state === 'downloading'
                    ? '下载中...'
                    : taskState[btn.type].state === 'processing'
                    ? '处理中...'
                    : taskState[btn.type].state === 'success'
                    ? '已下载'
                    : '下载文件'
                }}
              </button>
            </div>

          </template>

        </template>
        </div>
      </div>
    </div>
  </div><!-- /container -->
</template>