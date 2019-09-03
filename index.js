const express = require('express');
require('express-async-errors');
const session = require('express-session');
const app = express();
const axios = require('axios');
const querystring = require('querystring');
const bodyParser = require('body-parser');
const path = require("path");
app.use(bodyParser.urlencoded({
    extended: true
}));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(session({
    secret: 'oidwahuickl',
    resave: false,
    saveUninitialized: false
}));

app.get("/", async function (req, res) {
    const redirect_uri = `https://${req.get("host")}/token`;
    const ru = `https://developer.microsoft.com/en-us/graph/quick-start?appID=_appId_&appName=_appName_&redirectUrl=${redirect_uri}&platform=option-node`;
    const deepLink = `/quickstart/graphIO?publicClientSupport=false&appName=onefolder&redirectUrl=${redirect_uri}&allowImplicitFlow=false&ru=${encodeURIComponent(ru)}`;
    const app_url = "https://apps.dev.microsoft.com/?deepLink=" + encodeURIComponent(deepLink);
    res.render("index", { get_url: app_url });
});

app.post("/", async function (req, res) {
    req.session.client_id = req.body.client_id;
    req.session.client_secret = req.body.client_secret;
    res.redirect("/login");
});

app.get("/login", async function (req, res) {
    if (!req.session.client_id || !req.session.client_secret) {
        res.redirect(302, "/");
        return;
    }
    const client_id = req.session.client_id;
    const redirect_uri = `https://${req.get("host")}/token`;
    res.redirect(302, `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${client_id}&scope=${encodeURI("offline_access files.readwrite.all")}&response_type=code&redirect_uri=${redirect_uri}`);
});
app.get("/token", async function (req, res) {
    if (!req.session.client_id || !req.session.client_secret || !req.query.code) {
        res.status(400);
        res.send();
        return;
    }
    const redirect_uri = `https://${req.get("host")}/token`;
    const post_data = querystring.stringify({
        "client_id": req.session.client_id,
        "client_secret": req.session.client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
        "code": req.query.code
    });
    await axios.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", post_data, { headers: { "Content-Type": "application/x-www-form-urlencoded" } })
        .then(function (resp) {
            // handle success
            req.session.token = resp.data;
            req.session.token.client_id = req.session.client_id;
            req.session.token.client_secret = req.session.client_secret;
            req.session.token.redirect_uri = `https://${req.get("host")}/token`;
            res.render("token");
        })
        .catch(function (error) {
            // handle error
            console.log(error);
            res.send(error.response.data);
        });
});
app.get("/token.json", async function (req, res) {
    if (!req.session.token) {
        res.status(404).end();
        return;
    }
    res.set("Content-Disposition", "attachment; filename=\"token.json\"");
    res.send(req.session.token);
    req.session.destroy(function (err) {
        if (err) console.log(err);
    });
});
app.listen(3000, () => console.log("Run Success!"));