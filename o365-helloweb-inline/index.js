const util = require('util');
const exec = util.promisify(require('child_process').exec);

module.exports = async function (context, req) {

    if (!req.body && !req.body.webUrl) {
        context.res = {
            status: 400,
            body: "Please pass a webUrl in the request body"
        }
    }

    const login =
        { stdout, stderr }
        = await exec(`o365 login --authType password --userName '${process.env.USERNAME}' --password '${process.env.PASSWORD}'`);

    const web =
        { stdout, stderr }
        = await exec(`o365 spo web get --webUrl ${req.body.webUrl} --output json`)

    context.res = {
        body: web.stdout
    };

};