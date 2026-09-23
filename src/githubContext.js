const fs = require('fs')

// Minimal replacement for @actions/github's Context class. @actions/github
// ships as an ESM-only package starting with v3, which is incompatible with
// this project's CJS build (ncc) and test (jest) tooling, so we read the
// handful of fields this action actually needs directly from the same
// environment variables/files the GitHub Actions runner provides.
const loadContext = () => {
  const eventPath = process.env.GITHUB_EVENT_PATH
  let payload = {}
  if (eventPath && fs.existsSync(eventPath)) {
    payload = JSON.parse(fs.readFileSync(eventPath, { encoding: 'utf8' }))
  }
  return {
    payload,
    eventName: process.env.GITHUB_EVENT_NAME,
    workflow: process.env.GITHUB_WORKFLOW
  }
}

module.exports = { context: loadContext() }
