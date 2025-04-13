const { run } = require('./index');

// Mock the core module
jest.mock('@actions/core');
const core = require('@actions/core');

// Mock the process.env object
process.env = {};

const MSTeams = require('./MSTeams');


describe('run function', () => {
    const defaultParams = {
        webhook_url: 'dummy_webhook',
        job: JSON.stringify({ id: 'job_1', status: 'success' }),
        steps: JSON.stringify([{ name: 'step_1', status: 'success' }]),
        needs: JSON.stringify({ need_1: { result: 'success' } }),
        title: 'Test Title',
        msteams_emails: 'test@example.com',
        raw: '',
        dry_run: 'false',
    };

    beforeEach(() => {
        jest.clearAllMocks();
    });

    it('should send a notification with the correct payload', async () => {
        core.getInput.mockImplementation((name) => defaultParams[name]);
        const mockGeneratePayload = jest.spyOn(MSTeams.prototype, 'generatePayload').mockImplementation(jest.fn());
        
        mockGeneratePayload.mockResolvedValue({ message: 'payload' });

        await run();

        expect(mockGeneratePayload).toHaveBeenCalledTimes(1);

        expect(mockGeneratePayload).toHaveBeenCalledWith(expect.any(Object));
    });
 
    it('should handle invalid JSON in inputs gracefully', async () => {
        core.getInput.mockImplementation((name) => {
            if (name === 'job') return '{invalid_json';
            return '{}';
        });
        const mockSetFailed = jest.spyOn(core, 'setFailed');

        await run();

        expect(mockSetFailed).toHaveBeenCalled(            expect.string        );
    });

    it('should handle empty msteams_emails gracefully', async () => {
        const params = {
            ...defaultParams,
            msteams_emails: ''
        };
        core.getInput.mockImplementation((name) => params[name]);
        const mockNotify = jest.spyOn(MSTeams.prototype, 'notify').mockImplementation(jest.fn());

        await run();

        expect(mockNotify).toHaveBeenCalledWith('dummy_webhook', expect.any(Object));
    });

    it('should handle missing title gracefully', async () => {
        const params = {
            ...defaultParams,
            title: ''
        };
        core.getInput.mockImplementation((name) => params[name]);
        const mockNotify = jest.spyOn(MSTeams.prototype, 'notify').mockImplementation(jest.fn());

        await run();

        expect(mockNotify).toHaveBeenCalledWith('dummy_webhook', expect.any(Object));
    });

    it('should handle errors in generatePayload', async () => {
        const mockGeneratePayload = jest.spyOn(MSTeams.prototype, 'generatePayload').mockImplementation(jest.fn());
        mockGeneratePayload.mockRejectedValue(new Error('Payload generation failed'));

        const mockSetFailed = jest.spyOn(core, 'setFailed');

        await run();

        expect(mockSetFailed).toHaveBeenCalledWith('Payload generation failed');
    });

    it('should handle errors in notify', async () => {
        const mockNotify = jest.spyOn(MSTeams.prototype, 'notify').mockImplementation(jest.fn());
        mockNotify.mockRejectedValue(new Error('Notification failed'));

        const mockSetFailed = jest.spyOn(core, 'setFailed');

        await run();

        expect(mockSetFailed).toHaveBeenCalledWith(expect.any(String));
    });

    it('should allow raw payload', async () => {
        const params = {
            ...defaultParams,
            raw: JSON.stringify({ custom: 'payload' })
        };
        core.getInput.mockImplementation((name) => params[name]);
        const mockNotify = jest.spyOn(MSTeams.prototype, 'notify').mockImplementation(jest.fn());

        await run();

        expect(mockNotify).toHaveBeenCalledWith('dummy_webhook', { custom: 'payload' });
    });
});

describe('run function with dry_run', () => {
	beforeEach(() => {
		jest.clearAllMocks();
	});

	it('should skip notification when dry_run is true', async () => {
		core.getInput.mockImplementation((name) => {
			if (name === 'dry_run') return 'true';
			return '{}';
		});
		const mockNotify = jest.spyOn(MSTeams.prototype, 'notify').mockImplementation(jest.fn());
		const mockInfo = jest.spyOn(core, 'info');

		await run();

		expect(mockNotify).not.toHaveBeenCalled();
		expect(mockInfo).toHaveBeenCalledWith('Dry run - skipping notification send. Done.');
	});

	it('should send notification when dry_run is false', async () => {
		core.getInput.mockImplementation((name) => {
			if (name === 'dry_run') return 'false';
			if (name === 'webhook_url') return 'dummy_webhook';
			return '{}';
		});
		const mockNotify = jest.spyOn(MSTeams.prototype, 'notify').mockImplementation(jest.fn());
		const mockInfo = jest.spyOn(core, 'info');

		await run();

		expect(mockNotify).toHaveBeenCalled();
		expect(mockInfo).toHaveBeenCalledWith('Sent message to Microsoft Teams');
	});
});

describe('run function when webhook_url is missing', () => {
	beforeEach(() => {
		jest.clearAllMocks();
	});

	it('should throw an error when webhook_url is missing', async () => {
		core.getInput.mockImplementation((name) => {
			if (name === 'webhook_url') return '';
			return '{}';
		});
		const mockSetFailed = jest.spyOn(core, 'setFailed');

		await run();

		expect(mockSetFailed).toHaveBeenCalledWith(
			'[Error] Missing Microsoft Teams Incoming Webhooks URL.\n' +
			'Please configure "MSTEAMS_WEBHOOK" as environment variable or\n' +
			'specify the key called "webhook_url" in "with" section.'
		);
	});
});


describe('run function with various combinations of job, steps, and needs', () => {
    beforeEach(() => {
        jest.restoreAllMocks();
    });

    const defaults = {
        description: 'default',
        job: { id: 'job_1', status: 'success' },
        steps: [{ name: 'step_1', status: 'success', outputs: { key1: 'value1' } }],
        needs: { need_1: { result: 'success', outputs: { key2: 'value2' } } },
        msteams_emails: ''
    }

    // Define combinations of job, steps, and needs with outputs
    const combinations = [
        {
            // all parameters are present
            ...defaults
        },
        {
            // job is missing
            ...defaults,
            job: null
        },
        {
            // steps are empty
            ...defaults,
            steps: []
        },
        {
            // needs are empty
            ...defaults,
            needs: {}
        },
        {
            // all parameters are empty
            description: null,
            job: null,
            steps: [],
            needs: {},
        },
        {
            // steps contain a mix of statuses with outputs
            ...defaults,
            steps: [
                { name: 'step_1', status: 'success', outputs: { key1: 'value1' } },
                { name: 'step_2', status: 'failure', outputs: { key3: 'value3' } }
            ]
        },
        {
            // needs contains unsupported status
            ...defaults,
            needs: {
                need_1: { result: 'custom', outputs: { key2: 'value2' } },
                need_2: { result: 'failure', outputs: { key4: 'value4' } },
                // check truncation of needs ids
                abcdefghijklmnoprstuwxyz: { result: 'success', outputs: { key5: 'value5' } },
            },
            msteams_emails: 'foo1@bar.com, foo2@bar.com'
        }
    ];

    const testPayloadRemainsUnchanged = async ({ description, job, steps, needs, msteams_emails }) => {
        const mockNotify = jest.spyOn(MSTeams.prototype, 'notify').mockImplementation(jest.fn());
    
        // Mock inputs based on the combination
        core.getInput.mockImplementation((name) => {
            switch (name) {
                case 'job': return job ? JSON.stringify(job) : '{}';
                case 'steps': return steps ? JSON.stringify(steps) : '[]';
                case 'needs': return needs ? JSON.stringify(needs) : '{}';
                case 'webhook_url': return 'dummy_webhook';
                case 'title': return 'Test Title';
                case 'msteams_emails': return msteams_emails;
                default: return '';
            }
        });
    
        await run();
    
        // Verify MSTeams.notify is called with the correct payload
        expect(mockNotify).toHaveBeenCalledWith('dummy_webhook', expect.any(Object));
    
        // Use a snapshot to ensure the payload remains consistent
        expect(mockNotify.mock.calls[0][1]).toMatchSnapshot('payload');
    };

    it.each(combinations)('should call MSTeams.notify with correct payload when %j', testPayloadRemainsUnchanged);

    it('should handle missing inputs', async () => {
        core.getInput.mockImplementation((name) => undefined);
        
        const mockSetFailed = jest.spyOn(core, 'setFailed');

        await run();

        expect(mockSetFailed).toHaveBeenCalledWith(expect.any(String));
    });
});
