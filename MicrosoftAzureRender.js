var identification;
var config;

function MicrosoftAzure($)
{
    this.Info;
    this.Width;
    this.Height;
    this.Config;
    this.Welcome;

    // Databinding for property Info
    this.SetInfo = function(data)
    {
        ///UserCodeRegionStart:[SetInfo] (do not remove this comment.)
        this.Info = data;
        ///UserCodeRegionEnd: (do not remove this comment.)
    }

    // Databinding for property Info
    this.GetInfo = function()
    {
        ///UserCodeRegionStart:[GetInfo] (do not remove this comment.)
        return this.Info;
        ///UserCodeRegionEnd: (do not remove this comment.)
    }

    // Databinding for property Config
    this.SetConfig = function(data)
    {
        ///UserCodeRegionStart:[SetConfig] (do not remove this comment.)
        this.Config = data;
        ///UserCodeRegionEnd: (do not remove this comment.)
    }

    // Databinding for property Config
    this.GetConfig = function()
    {
        ///UserCodeRegionStart:[GetConfig] (do not remove this comment.)
        return this.Config;
        ///UserCodeRegionEnd: (do not remove this comment.)
    }

    this.show = function()
    {
        ///UserCodeRegionStart:[show] (do not remove this comment.)

        var buffer   = "";
        var _this    = this;
        var BasePath = gx.util.resourceUrl(gx.basePath + gx.staticDirectory + "MicrosoftAzure/", true);

        if (!this.IsPostBack) {

            buffer += '<h4 id="WelcomeMessage"></h4>'
            buffer += '<button id="SignIn" class="SignIn" onclick="' + this.me() + '.SignIn(); return false">Sign In'
            buffer += '   <img class="btnimg" src="'+ BasePath + '/sigin.png" alt="" />'
            buffer += '</button>'
        }

        $('#' + this.ContainerName).html(buffer);

        this.SignIn = function() {

            myMSALObj.loginPopup(applicationConfig.graphScopes).then(function(idToken) {
                showWelcomeMessage();
                acquireTokenPopupAndCallMSGraph();
                _this.Info = identification;
                _this.Sign();
            }, function(error) {
                console.log(error);
            });
        }

        this.SignOut = function(){
            myMSALObj.logout();
            _this.Quit();
        }


        var applicationConfig = {
            clientID: this.Config.clientID,
            authority: this.Config.authority,
            graphScopes: JSON.parse("[" + this.Config.graphScopes + "]"), //["user.read"],
            graphEndpoint: this.Config.graphEndpoint
        };

        var myMSALObj = new Msal.UserAgentApplication(applicationConfig.clientID, null, acquireTokenRedirectCallBack, {
            storeAuthStateInCookie: true,
            cacheLocation: "localStorage"
        });


        function acquireTokenPopupAndCallMSGraph() {
            myMSALObj.acquireTokenSilent(applicationConfig.graphScopes).then(function(accessToken) {
                callMSGraph(applicationConfig.graphEndpoint, accessToken, graphAPICallback);
            }, function(error) {
                console.log(error);                
                if (error.indexOf("consent_required") !== -1 || error.indexOf("interaction_required") !== -1 || error.indexOf("login_required") !== -1) {
                    myMSALObj.acquireTokenPopup(applicationConfig.graphScopes).then(function(accessToken) {
                        callMSGraph(applicationConfig.graphEndpoint, accessToken, graphAPICallback);
                    }, function(error) {
                        console.log(error);
                    });
                }
            });
        }

        function graphAPICallback(data) {
            console.log(JSON.stringify(data, null, 2));
        }

        function showWelcomeMessage() {
            identification = myMSALObj.getUser();
            var divWelcome = document.getElementById('WelcomeMessage');
            divWelcome.innerHTML += _this.Welcome + ' ' +  myMSALObj.getUser().name;
            var loginbutton = document.getElementById('SignIn');            
            loginbutton.innerHTML = 'Sign Out';
            loginbutton.className = 'SignOut';
            loginbutton.setAttribute('onclick',  _this.me() + '.SignOut();');
                    
        }
        
        function acquireTokenRedirectAndCallMSGraph() {        
            myMSALObj.acquireTokenSilent(applicationConfig.graphScopes).then(function(accessToken) {
                callMSGraph(applicationConfig.graphEndpoint, accessToken, graphAPICallback);
            }, function(error) {
                console.log(error);
                if (error.indexOf("consent_required") !== -1 || error.indexOf("interaction_required") !== -1 || error.indexOf("login_required") !== -1) {
                    myMSALObj.acquireTokenRedirect(applicationConfig.graphScopes);
                }
            });
        }

        function acquireTokenRedirectCallBack(errorDesc, token, error, tokenType) {
            if (tokenType === "access_token") {
                callMSGraph(applicationConfig.graphEndpoint, token, graphAPICallback);
            } else {
                console.log("token type is:" + tokenType);
            }
        }

        var ua = window.navigator.userAgent;
        var msie = ua.indexOf('MSIE ');
        var msie11 = ua.indexOf('Trident/');
        var msedge = ua.indexOf('Edge/');
        var isIE = msie > 0 || msie11 > 0;
        var isEdge = msedge > 0;

        if (!isIE) {
            if (myMSALObj.getUser()) { 
                showWelcomeMessage();
                acquireTokenPopupAndCallMSGraph();
            }
        } else {
            document.getElementById("SignIn").onclick = function() {
                myMSALObj.loginRedirect(applicationConfig.graphScopes);
            };
            if (myMSALObj.getUser() && !myMSALObj.isCallback(window.location.hash)) { 
                showWelcomeMessage();
                acquireTokenRedirectAndCallMSGraph();
            }
        }        
        
        ///UserCodeRegionEnd: (do not remove this comment.)
    }
    ///UserCodeRegionStart:[User Functions] (do not remove this comment.)
    ///UserCodeRegionEnd: (do not remove this comment.):
}
