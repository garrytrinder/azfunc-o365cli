const util = require('util');
const exec = util.promisify(require('child_process').exec);

module.exports = async function (context, req) {

    if (!req.body && !req.body.webUrl) {
        context.res = {
            status: 400,
            body: "Please pass a webUrl in the request body"
        }
    }

    const cliPath = `${__dirname}/../node_modules/.bin/o365`;

    const login =
        { stdout, stderr }
        = await exec(`${cliPath} login --authType password --userName '${process.env.SERVICEACCOUNT_NAME}' --password '${process.env.SERVICEACCOUNT_PASSWORD}'`);

    const web =
        { stdout, stderr }
        = await exec(`${cliPath} spo web get --webUrl ${req.body.webUrl} --output json`)

    context.res = {
        body: web.stdout
    };

};