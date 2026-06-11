import { queryExportTask } from '@/api/index'

export function usePolling() {
  const timers = {}

  const startPolling = (
    type,
    token,
    onSuccess,
    onError
  ) => {
    const poll = async () => {
      try {
        const res = await queryExportTask(token)

        console.log(`[poll] ${type}`, res.data.blocks.resultBlock.data)

        onSuccess?.(res)
      } catch (err) {
        console.error(err)

        onError?.(err)
      }
    }

    poll()

    timers[type] = setInterval(
      poll,
      10000
    )
  }

  const stopPolling = (type) => {
    if (timers[type]) {
      clearInterval(timers[type])
      delete timers[type]
    }
  }

  return {
    startPolling,
    stopPolling
  }
}