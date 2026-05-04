const core = require('@actions/core');
const MSTeams = require('./MSTeams');

const missing_functionality_warning = objective =>
	core.warning(`Missing ${objective} parameter will result in reduced functionality.`) || {};

const access_context = context_name => {
	let context = core.getInput(context_name);
	if (!context) missing_functionality_warning(context_name);
	return context === '' ? {} : JSON.parse(context);
};

async function run() {
	try {
		const webhook_url = process.env.MSTEAMS_WEBHOOK || core.getInput('webhook_url');
		if (webhook_url === '') {
			throw new Error(
				'[Error] Missing Microsoft Teams Incoming Webhooks URL.\n' +
				'Please configure "MSTEAMS_WEBHOOK" as environment variable or\n' +
				'specify the key called "webhook_url" in "with" section.'
			);
		}

		// Register webhook_url as a secret so it is masked in logs
		core.setSecret(webhook_url);

		// Parse secrets input and register each value for masking
		const secretsInput = core.getInput('secrets');
		let secretValues = [];
		if (secretsInput) {
			try {
				const secrets = JSON.parse(secretsInput);
				secretValues = Object.values(secrets).filter(v => typeof v === 'string' && v.length > 0);
				secretValues.forEach(value => core.setSecret(value));
			} catch (e) {
				core.warning('Failed to parse secrets input. Secret masking may not be fully applied.');
			}
		}

		// Replace all known secret values in a string with '***'
		const maskSecrets = (str) => {
			if (!secretValues.length || typeof str !== 'string') return str;
			let result = str;
			for (const secret of secretValues) {
				result = result.split(secret).join('***');
			}
			return result;
		};

		let job = access_context('job');
		let steps = access_context('steps');
		let needs = access_context('needs');

		let title = core.getInput('title');
		let msteams_emails= core.getInput('msteams_emails');
		let raw = core.getInput('raw');
		let dry_run = core.getInput('dry_run');

		const overwrite = core.getInput('overwrite');
		if (overwrite) {
			core.warning('The "overwrite" parameter is deprecated. Please use "raw" instead.');
			if (!raw) raw = overwrite;
		}

		core.info(`Parsed params:\n${JSON.stringify({
			webhook_url: '***',
			job,
			steps,
			needs,
			raw,
			title,
			msteams_emails,
			dry_run
		})}`);

		const msteams = new MSTeams();
		let payload;
		if (raw === '') {
			payload = await msteams.generatePayload(
				{
					job,
					steps,
					needs,
					title,
					msteams_emails
				}
			);
		} else {
			payload = Object.assign({}, msteams.header, JSON.parse(raw));
		}

		// Mask any secret values present in the payload before sending or logging
		payload = JSON.parse(maskSecrets(JSON.stringify(payload)));

    try {
      core.info(
        `Generated payload for Microsoft Teams:\n${JSON.stringify(
          payload,
          null,
          2
        )}`
      );
    } catch (stringifyError) {
      core.error(
        `Generated payload for Microsoft Teams (contains circular references, showing keys only):
		${stringifyError}`
      );
    }

		if (dry_run === '' || dry_run==='false') {
			await msteams.notify(webhook_url, payload);
			core.info('Sent message to Microsoft Teams');
		} else {
			core.info('Dry run - skipping notification send. Done.');
		}
	} catch (err) {
		core.setFailed(err.message);
	}
}

if (require.main === module) {
	run();
} else {
	exports.run = run;
}
