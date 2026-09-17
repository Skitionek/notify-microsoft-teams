// src/MSTeams.test.js
const MSTeams = require('./MSTeams')
const http = require('@actions/http-client')

// Mock the github context
jest.mock('@actions/github', () => ({
  context: {
    payload: {
      repository: {
        html_url: 'html_url',
        name: 'name'
      },
      compare: 'compare_url',
      sender: {
        login: 'login',
        url: 'url'
      },
      commits: [],
      head_commit: {
        timestamp: 'timestamp'
      }
    },
    eventName: 'push',
    workflow: 'test_workflow'
  }
}))

jest.mock('@actions/http-client')

describe('MSTeams.generatePayload', () => {
  it('should use custom actions when provided', async () => {
    const customActions = [
      { type: 'Action.OpenUrl', title: 'Custom Link', url: 'https://example.com' }
    ]
    const msTeams = new MSTeams()
    const payload = await msTeams.generatePayload({ actions: customActions })

    const actionSet = payload.attachments[0].content.body.find(b => b.type === 'ActionSet')
    expect(actionSet).toBeDefined()
    expect(actionSet.actions).toEqual(customActions)
  })

  it('should use default Repository/Compare actions when actions is null', async () => {
    const msTeams = new MSTeams()
    const payload = await msTeams.generatePayload({ actions: null })

    const actionSet = payload.attachments[0].content.body.find(b => b.type === 'ActionSet')
    expect(actionSet).toBeDefined()
    expect(actionSet.actions).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ title: 'Repository' }),
        expect.objectContaining({ title: 'Compare' })
      ])
    )
  })

  it('should use default Repository/Compare actions when actions is not provided', async () => {
    const msTeams = new MSTeams()
    const payload = await msTeams.generatePayload({})

    const actionSet = payload.attachments[0].content.body.find(b => b.type === 'ActionSet')
    expect(actionSet).toBeDefined()
    expect(actionSet.actions.some(a => a.title === 'Repository')).toBe(true)
  })
})

describe('MSTeams.notify', () => {
  const webhookUrl = 'test-webhook-url'
  const payload = { message: 'Test Payload' }

  let mockPostJson

  beforeEach(() => {
    jest.clearAllMocks()
    jest.restoreAllMocks()

    // Mock the HttpClient class and its postJson method
    mockPostJson = jest.fn()
    http.HttpClient.mockImplementation(() => ({
      postJson: mockPostJson
    }))
  })

  it('should send a success notification with status 202', async () => {
    mockPostJson.mockResolvedValueOnce({ statusCode: 202 })

    const msTeams = new MSTeams()
    await msTeams.notify(webhookUrl, payload)

    expect(http.HttpClient).toHaveBeenCalledWith()
    expect(mockPostJson).toHaveBeenCalledWith(webhookUrl, payload)
    expect(mockPostJson).toHaveBeenCalledTimes(1)
  })

  it('should send a success notification with status 200', async () => {
    mockPostJson.mockResolvedValueOnce({ statusCode: 200 })

    const msTeams = new MSTeams()
    await msTeams.notify(webhookUrl, payload)

    expect(http.HttpClient).toHaveBeenCalledWith()
    expect(mockPostJson).toHaveBeenCalledWith(webhookUrl, payload)
    expect(mockPostJson).toHaveBeenCalledTimes(1)
  })

  it('should throw an error if the notification fails', async () => {
    mockPostJson.mockRejectedValueOnce(new Error('Webhook error'))

    const msTeams = new MSTeams()
    await expect(msTeams.notify(webhookUrl, payload)).rejects.toThrow(
      expect.any(Error)
    )

    expect(http.HttpClient).toHaveBeenCalledWith()
    expect(mockPostJson).toHaveBeenCalledWith(webhookUrl, payload)
    expect(mockPostJson).toHaveBeenCalledTimes(1)
  })

  it('should throw an error for missing webhookUrl', async () => {
    const msTeams = new MSTeams()
    await expect(msTeams.notify(undefined, payload)).rejects.toThrow(
      expect.any(Error)
    )

    expect(http.HttpClient).not.toHaveBeenCalled()
    expect(mockPostJson).not.toHaveBeenCalled()
  })

  it('should throw an error for missing payload', async () => {
    const msTeams = new MSTeams()
    await expect(msTeams.notify(webhookUrl, undefined)).rejects.toThrow(
      expect.any(Error)
    )

    expect(http.HttpClient).not.toHaveBeenCalled()
    expect(mockPostJson).not.toHaveBeenCalled()
  })

  it('Returns error for empty response', async () => {
    mockPostJson.mockResolvedValueOnce({})

    const msTeams = new MSTeams()
    await expect(msTeams.notify(webhookUrl, payload)).rejects.toThrow(
      expect.any(Error)
    )

    expect(http.HttpClient).toHaveBeenCalledWith()
    expect(mockPostJson).toHaveBeenCalledWith(webhookUrl, payload)
    expect(mockPostJson).toHaveBeenCalledTimes(1)
  })

  it('Handles response with circular references without JSON.stringify errors', async () => {
    // Create a mock response object with circular references
    const mockResult = { error: 'Invalid payload' }

    // Create circular reference
    const circular = { ref: mockResult }
    mockResult.circular = circular

    const mockResponseWithCircularRef = {
      statusCode: 400,
      result: mockResult,
      headers: { 'content-type': 'application/json' }
    }

    // This creates the circular reference that would cause JSON.stringify to fail
    mockPostJson.mockResolvedValueOnce(mockResponseWithCircularRef)

    const msTeams = new MSTeams()

    // This should throw an error but NOT a circular reference error
    try {
      await msTeams.notify(webhookUrl, payload)
      // Should not reach here
      expect(true).toBe(false)
    } catch (error) {
      // Verify the error message contains the safe response data and not circular reference errors
      expect(error.message).toContain(
        'Failed to send notification to Microsoft Teams'
      )
      expect(error.message).toContain('"statusCode": 400')
      expect(error.message).not.toContain(
        'Converting circular structure to JSON'
      )

      // Verify that the error message includes the response details we expect
      const errorMessage = error.message
      expect(errorMessage).toMatch(/"statusCode":\s*400/)
    }

    expect(http.HttpClient).toHaveBeenCalledWith()
    expect(mockPostJson).toHaveBeenCalledWith(webhookUrl, payload)
  })
})
