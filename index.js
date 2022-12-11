const express = require('express')
require('dotenv').config()
const app = express()
const port = process.env.PORT
const passport = require('passport');
const InfusionsoftStrategy  = require('@albertopronails/passport-infusionsoft').Strategy;
const axios = require('axios');
const bodyParser = require('body-parser');
app.use(passport.initialize());
app.use(passport.session());
app.use(bodyParser.json({limit: '50mb'}));
app.use(bodyParser.urlencoded({limit: '50mb', extended: true}));
//require('./allcron');
const myCache = require('./myCache');
const dataProduct = require('./getProduct')


app.get('/', (req, res) => {
    // todo: htpassword nginx docker
    res.sendFile(__dirname+'/static/index.html')

})

app.listen(port, () => {
    console.log(`Demo http://localhost:${port}`)
})

passport.use(new InfusionsoftStrategy({
        clientID: process.env.CLIENTID,
        clientSecret: process.env.CLIENTSECRET,
        callbackURL: process.env.CALLBACKURL
    },
    function(accessToken, refreshToken, profile, done) {



            const tokens = {
                accessToken: accessToken,
                refreshToken: refreshToken
            };
            myCache.set( "tokens", tokens, 0 );
            console.log(tokens)

             dataProduct.createDatabaseOffice()

            return done(null, tokens);
    }
));

passport.serializeUser(function(user, done) {
    done(null, JSON.stringify(user));
});

passport.deserializeUser(function(user, done) {
    done(null, JSON.parse(user));
});

app.get('/auth/infusionsoft',
    passport.authenticate('Infusionsoft'),
    function(req, res){
    });

app.get('/auth/infusionsoft/callback',
    passport.authenticate('Infusionsoft', { failureRedirect: '/login' }),
    function(req, res) {
        res.redirect('/');
    });



