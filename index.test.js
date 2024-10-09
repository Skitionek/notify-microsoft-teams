const core = require('@actions/core');

const action = './src';

const webhook_url = process.env.MSTEAMS_WEBHOOK;
const dry_run = process.env.DRY_RUN.toLowerCase() == "true";

const setInputsMock = (inputs) => 
    jest.fn().mockImplementation((name, options) => inputs[name]);

afterEach(() => {
    jest.resetModules();
});

test('One with everything', () => {
    core.getInput = setInputsMock({
        webhook_url,
        needs: JSON.stringify({

        }),
        job: JSON.stringify({
            
        }),
        steps: JSON.stringify({
            
        }),
        dry_run
    });
    require(action);
});

test('One with needs', () => {
    core.getInput = setInputsMock({
        webhook_url,
        needs: JSON.stringify({

        }),
        job: JSON.stringify({
            
        }),
        steps: JSON.stringify({
            
        }),
        dry_run
    });
    require('./src');
});

test('One with steps', () => {
    core.getInput = setInputsMock({
        webhook_url,
        needs: JSON.stringify({

        }),
        job: JSON.stringify({
            
        }),
        steps: JSON.stringify({
            
        }),
        dry_run
    });
    require(action);
});

test('One failure', () => {
    core.getInput = setInputsMock({
        webhook_url,
        needs: JSON.stringify({

        }),
        job: JSON.stringify({
            
        }),
        steps: JSON.stringify({
            
        }),
        dry_run
    });
    require(action);
});

test('One with little info', () => {
    core.getInput = setInputsMock({
        webhook_url,
        dry_run
    });
    require(action);
});

test('One without emails', () => {
    core.getInput = setInputsMock({
        webhook_url,
        title: "`Overwrote title in ${workflow_link}`",
        job: JSON.stringify({
            
        }),
        steps: JSON.stringify({
            
        }),
        dry_run
    });
    require(action);
});

test('One with emails', () => {
    core.getInput = setInputsMock({
        webhook_url,
        title: "`Overwrote title in ${workflow_link}`",
        msteams_emails: "mm@mm.mm, yy@yy.yy, rr@rr.rr",
        job: JSON.stringify({
            
        }),
        steps: JSON.stringify({
            
        }),
        dry_run
    });
    require(action);
});

test('One with overwrite', () => {
    core.getInput = setInputsMock({
        webhook_url,
        title: "`Overwrote title in ${workflow_link}`",
        job: JSON.stringify({
            
        }),
        steps: JSON.stringify({
            
        }),
        dry_run
    });
    require(action);
});

test('One with raw data', () => {
    core.getInput = setInputsMock({
        webhook_url,
        raw: JSON.stringify({
            "type": "message",
            "attachments":
            [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content":
                    {
                        "type": "AdaptiveCard",
                        "body":
                        [
                            {
                                "type": "TextBlock",
                                "size": "Medium",
                                "weight": "Bolder",
                                "text": "Test title text"
                            }
                        ]
                    }
                }
            ]
        }),
        dry_run
    });
    require(action);
});
