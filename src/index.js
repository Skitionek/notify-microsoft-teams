const core = require('@actions/core')
const MSTeams = require('./MSTeams')

const missing_functionality_warning = objective =>
  core.warning(
    `Missing ${objective} parameter will result in reduced functionality.`
  ) || {}

const isDebugFlagEnabled = value => {
  if (!value) return false

  const normalizedValue = String(value).trim().toLowerCase()
  return normalizedValue === '1' || normalizedValue === 'true'
}

const isDebugEnabled = () =>
  isDebugFlagEnabled(process.env.RUNNER_DEBUG) ||
  isDebugFlagEnabled(process.env.ACTIONS_STEP_DEBUG)

const serializeError = err => {
  if (!err) return 'Unknown error'
  if (err.stack) return err.stack
  if (err.message) return err.message

  try {
    return JSON.stringify(err, null, 2)
  } catch {
    return String(err)
  }
}

const logError = (message, err) => {
  core.error(`${message}: ${err && err.message ? err.message : String(err)}`)
  if (isDebugEnabled()) {
    core.debug(`Error details:\n${serializeError(err)}`)
  }
}

const parseRetryInput = value => {
  if (!value) return 0

  const retries = Number.parseInt(value, 10)
  if (Number.isNaN(retries) || retries < 0) {
    throw new Error(
      `Invalid "retries" input: "${value}". Please provide a non-negative integer.`
    )
  }

  return retries
}

const access_context = context_name => {
  const context = core.getInput(context_name)
  if (!context) missing_functionality_warning(context_name)
  return context === '' ? {} : JSON.parse(context)
}

async function run () {
  try {
    const webhook_url =
      process.env.MSTEAMS_WEBHOOK || core.getInput('webhook_url')
    if (webhook_url === '') {
      throw new Error(
        '[Error] Missing Microsoft Teams Incoming Webhooks URL.\n' +
          'Please configure "MSTEAMS_WEBHOOK" as environment variable or\n' +
          'specify the key called "webhook_url" in "with" section.'
      )
    }

    const job = access_context('job')
    const steps = access_context('steps')
    const needs = access_context('needs')

    const title = core.getInput('title')
    const actions = core.getInput('actions')
    const msteams_emails = core.getInput('msteams_emails')
    let raw = core.getInput('raw')
    const dry_run = core.getInput('dry_run')
    const retries = parseRetryInput(core.getInput('retries'))

    const overwrite = core.getInput('overwrite')
    if (overwrite) {
      core.warning(
        'The "overwrite" parameter is deprecated. Please use "raw" instead.'
      )
      if (!raw) raw = overwrite
    }

    core.info(
      `Parsed params:\n${JSON.stringify({
        webhook_url: '***',
        job,
        steps,
        needs,
        raw,
        title,
        actions,
        msteams_emails,
        dry_run,
        retries
      })}`
    )

    if (isDebugEnabled()) {
      core.debug('GitHub Actions step debug logging is enabled.')
    }

    const msteams = new MSTeams()
    let payload
    if (raw === '') {
      let parsedActions = null
      if (actions) {
        try {
          parsedActions = JSON.parse(actions)
        } catch (e) {
          throw new Error(
            `Invalid JSON provided for "actions" input: ${e.message}. Please ensure the "actions" input is a valid JSON array of Adaptive Card Action objects (see https://adaptivecards.io/explorer/Action.OpenUrl.html).`
          )
        }
      }
      payload = await msteams.generatePayload({
        job,
        steps,
        needs,
        title,
        actions: parsedActions,
        msteams_emails
      })
    } else {
      payload = JSON.parse(raw)
    }

    try {
      core.info(
        `Generated payload for Microsoft Teams:\n${JSON.stringify(
          payload,
          null,
          2
        )}`
      )
    } catch (stringifyError) {
      core.error(
        `Generated payload for Microsoft Teams (contains circular references, showing keys only):
		${stringifyError}`
      )
    }

    if (dry_run === '' || dry_run === 'false') {
      const attempts = retries + 1
      let sent = false
      let lastError

      for (let attempt = 1; attempt <= attempts; attempt++) {
        if (isDebugEnabled()) {
          core.debug(`Sending notification attempt ${attempt}/${attempts}`)
        }

        try {
          await msteams.notify(webhook_url, payload)
          sent = true
          core.info('Sent message to Microsoft Teams')
          break
        } catch (err) {
          lastError = err
          logError(`Notification attempt ${attempt}/${attempts} failed`, err)

          if (attempt < attempts) {
            core.warning(`Retrying Microsoft Teams notification (${attempt + 1}/${attempts})`)
          }
        }
      }

      if (!sent) {
        const lastErrorMessage =
          lastError && lastError.message
            ? lastError.message
            : serializeError(lastError)

        throw new Error(
          `Failed to send notification to Microsoft Teams after ${attempts} attempt(s): ${lastErrorMessage}`
        )
      }
    } else {
      core.info('Dry run - skipping notification send. Done.')
    }
  } catch (err) {
    logError('Action execution failed', err)
    const failureMessage = err && err.message ? err.message : serializeError(err)
    core.setFailed(failureMessage)
  }
}

if (require.main === module) {
  run()
} else {
  exports.run = run
}
