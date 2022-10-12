const {IncomingWebhook} = require('ms-teams-webhook');
const {context: github} = require('@actions/github');
const merge = require('lodash.merge');
const core = require('@actions/core');

const placeholder = '';
const {
    payload: {
        repository = {
            html_url: placeholder,
            name: placeholder
        },
        compare,
        sender = {
            login: placeholder,
            url: placeholder
        },
        commits = [],
        head_commit,
        pull_request
    },
    eventName,
    workflow,
    sha
} = github;

const timestamp = head_commit ? head_commit.timestamp : pull_request ? pull_request.updated_at : placeholder;

const statuses = [
    {
        id: 'success',
        icon: '✓',
        color: '#2cbe4e',
        activityTitle: "Success!",
        activitySubtitle: timestamp,
        activityImage: "https://raw.githubusercontent.com/Skitionek/notify-microsoft-teams/master/icons/success.png"

    },
    {
        id: 'failure',
        icon: '✗',
        color: '#cb2431',
        activityTitle: "Failure",
        activitySubtitle: timestamp,
        activityImage: "https://raw.githubusercontent.com/Skitionek/notify-microsoft-teams/master/icons/failure.png"

    },
    {
        id: 'cancelled',
        icon: 'o',
        color: '#ffc107',
        activityTitle: "Cancelled",
        activitySubtitle: timestamp,
        activityImage: "https://raw.githubusercontent.com/Skitionek/notify-microsoft-teams/master/icons/cancelled.png"
    },
    {
        id: 'skipped',
        icon: '⤼',
        color: '#1a6aff',
        activityTitle: "Skipped",
        activitySubtitle: timestamp,
        activityImage: "https://raw.githubusercontent.com/Skitionek/notify-microsoft-teams/master/icons/skipped.png"
    },
    {
        id: 'unknown',
        icon: '?',
        color: '#999',
        activityTitle: 'No job context has been provided',
        activitySubtitle: timestamp,
        activityImage: "https://raw.githubusercontent.com/Skitionek/notify-microsoft-teams/master/icons/unknown.png"
    }
];

function Status(status) {
    if (!status) {
        core.error(`Unknown status value: ${status}`);
        return statuses.find(({id}) => id === 'unknown')
    }
    const r = statuses.find(({id}) => id === status.toLowerCase());
    if (!r) {
        core.error(`Not implemented status value: ${status}`)
        return statuses.find(({id}) => id === 'unknown')
    }
    return r
}

function oneLineFromMultilineCommitMsg(msg) {
    return msg.split(/\n/g).reduce((o, c) => o + (o.length > 0 && c.trim().length > 0 && o[o.length - 2] != ',' && o[o.length - 1] != ' ' ? ", " : '') + c.trim(), '')
}

const repository_link = `[${repository.full_name}](${repository.html_url})`;
const pr_link = pull_request ? `${pull_request.html_url}` : undefined;
const changelog = commits.length ? `${commits.reduce((o, c) => o + '\n+ ' + oneLineFromMultilineCommitMsg(c.message), placeholder)}` : undefined;
const title_text = head_commit ? `[Push] ${head_commit.message.replace(/\n.*/g, '')}` : pull_request ? `[PR] ${pull_request.title}` : `${eventName == 'workflow_dispatch' ? 'Manually' : eventName} triggered "${workflow}"`;
const outputs2markdown = (outputs) =>
    Object.keys(outputs).reduce((o, output_name) => o + `+ ${output_name}:${'\n'}\`\`\`${outputs[output_name]}\`\`\``, '');

const summary_generator = (obj, status_key) => {
    const r = {
        facts: [],
        text: '',
        startGroup: true
    };
    Object.keys(obj).forEach(step_id => {
        const status = Status(obj[step_id][status_key]);
        r.facts.push({
            name: `${status.icon} ${step_id}`,
            value: status.activityTitle
        });
        if (status.id === 'failure' && obj[step_id].outputs.length) {
            r.text += `${step_id}:\n`;
            r.text += outputs2markdown(obj[step_id].outputs)
        }
    });
    if (r.text === '') delete r.text;
    if (!r.facts.length) return [];
    return [r]
};

class MSTeams {
    constructor() {
        this.header = {
            "@type": "MessageCard",
            "@context": "http://schema.org/extensions"
        }
    }

    /**
     * Generate msteams payload
     * @param {string} jobName
     * @returns
     */
    async generatePayload(
        {
            job = {status: 'unknown'},
            steps = {},
            needs = {},
            overwrite = '',
            run_id = undefined
        }
    ) {
        const steps_summary = summary_generator(steps, 'outcome');
        const needs_summary = summary_generator(needs, 'result');

        const {
            activityTitle,
            activitySubtitle,
            activityImage,
            color
        } = Status(job.status);
        const status_summary = {
            activityTitle,
            activitySubtitle,
            activityImage
        };

        const sections = [
            ...steps_summary,
            ...needs_summary,
            status_summary
        ];
        const payload = {
            ...this.header,
            correlationId: sha,
            themeColor: color,
            title: title_text,
            summary: repository_link,
            sections,
            text: `by **${sender.login}** on **${repository.name}**`,
            potentialAction: [
                {
                    "@type": "OpenUri",
                    name: "Repository",
                    targets: [
                        {os: "default", uri: repository.html_url}
                    ]
                },
                {
                    "@type": "OpenUri",
                    name: "Pull Request",
                    targets: [
                        {os: "default", uri: pr_link}
                    ]
                },
                {
                    "@type": "OpenUri",
                    name: "Workflow Run",
                    targets: [
                        {
                            os: "default",
                            uri: run_id ? `${repository.html_url}/actions/runs/${run_id}` : `${repository.html_url}/actions?query=workflow%3A${workflow}`
                        }
                    ]
                },
                {
                    "@type": "OpenUri",
                    name: "Compare",
                    targets: [
                        {os: "default", uri: compare}
                    ]
                }
            ]
        };
        if (changelog) {
            const changelog_summary = {
                facts: [
                    {
                        name: "Changelog",
                        value: changelog
                    }
                ]
            }
            payload.sections.push(changelog_summary)
        }

        if (overwrite !== '') {
            return merge(
                payload,
                eval(`(${overwrite})`)
            )
        } else {
            return payload
        }
    }

    /**
     * Notify information about github actions to MSTeams
     * @param url
     * @param  payload
     * @returns {Promise} result
     */
    async notify(
        url,
        payload
    ) {
        const client = new IncomingWebhook(url);
        const response = await client.send(payload);

        if (!response.text) {
            throw new Error(
                "Failed to send notification to Microsoft Teams.\n" +
                "Response:\n" +
                JSON.stringify(response, null, 2)
            );
        }
    }
}

module.exports = MSTeams;