// src/MSTeams.test.js
const MSTeams = require('./MSTeams');
const { IncomingWebhook } = require('ms-teams-webhook');

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
}));

jest.mock('ms-teams-webhook');

describe('MSTeams.notify', () => {
  const webhookUrl = 'test-webhook-url';
  const payload = { message: 'Test Payload' };

  let mockSend;

  beforeEach(() => {
    jest.clearAllMocks();
    jest.restoreAllMocks();

    // Mock the IncomingWebhook class and its send method
    mockSend = jest.fn();
    IncomingWebhook.mockImplementation(() => ({
      send: mockSend,
    }));
  });

  it('should send a success notification', async () => {
    mockSend.mockResolvedValueOnce({ text: 'ok' });

    const msTeams = new MSTeams();
    await msTeams.notify(webhookUrl, payload);

    expect(IncomingWebhook).toHaveBeenCalledWith(webhookUrl);
    expect(mockSend).toHaveBeenCalledWith(payload);
    expect(mockSend).toHaveBeenCalledTimes(1);
  });

  it('should throw an error if the notification fails', async () => {
    mockSend.mockRejectedValueOnce(new Error('Webhook error'));

    const msTeams = new MSTeams();
    await expect(msTeams.notify(webhookUrl, payload)).rejects.toThrow(expect.any(Error));

    expect(IncomingWebhook).toHaveBeenCalledWith(webhookUrl);
    expect(mockSend).toHaveBeenCalledWith(payload);
    expect(mockSend).toHaveBeenCalledTimes(1);
  });

  it('should throw an error for missing webhookUrl', async () => {
    const msTeams = new MSTeams();
    await expect(msTeams.notify(undefined, payload)).rejects.toThrow(expect.any(Error));

    expect(IncomingWebhook).not.toHaveBeenCalled();
    expect(mockSend).not.toHaveBeenCalled();
  });

  it('should throw an error for missing payload', async () => {
    const msTeams = new MSTeams();
    await expect(msTeams.notify(webhookUrl, undefined)).rejects.toThrow(expect.any(Error));

    expect(IncomingWebhook).not.toHaveBeenCalled();
    expect(IncomingWebhook.prototype.send).not.toHaveBeenCalled();
  });

  it('Returns error for empty response', async () => {
    mockSend.mockResolvedValueOnce({});

    const msTeams = new MSTeams();
    await expect(msTeams.notify(webhookUrl, payload)).rejects.toThrow(expect.any(Error));

    expect(IncomingWebhook).toHaveBeenCalledWith(webhookUrl);
    expect(mockSend).toHaveBeenCalledWith(payload);
    expect(mockSend).toHaveBeenCalledTimes(1);
  });
});