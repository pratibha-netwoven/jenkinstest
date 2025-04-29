    package com.example.teamsSend
    import groovy.json.JsonBuilder
    import groovy.json.JsonSlurper

    class MsTeamsHelper {

      /** Azure AD App Details*/
        static String TENANT_ID = '********-****-****-****-************'
        static String CLIENT_ID = '********-****-****-****-************'
        static String CLIENT_SECRET = 'L9c0***********************Tz8='
        /* Adjust based on required permissions */
        static String SCOPE = "https://service.flow.microsoft.com//.default" /* do not try to fix the URL. "//" is required to generate the required access token*/
                                                                            // "https://graph.microsoft.com/.default"
        static String  TOKEN_URL ="https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token"

    

     /**
        * Sends a formatted Adaptive Card message to a Microsoft Teams channel via webhook anonymous.
        *
        * Usage:
        * teamsSend(
        *     teamsWebhookUrl,     // (Required) Teams webhook URL.
        *     type,                // (Required) 'post' or 'reply'.
        *     teamsTeamName,       // (Required) Title of the Teams team.
        *     teamsChannelName,    // (Required) Title of the Teams channel.
        *     threadId,            // (Optional) Leave blank for new root post.
        *                          //           Provide threadId to update a root post or reply to an existing post.
        *     replyId,             // (Optional) Leave blank for new reply.
        *                          //           Provide replyId to update an existing reply.
        *     status,              // (Required) Message status: 'loading', 'success', or 'failure'.
        *     msgTitle,            // (Required) Title of the post or reply.
        *     msgBody              // (Required) Body content of the post or reply.
        * )
        *
        * Workflow:
        * 1. Builds the message payload with status/type-specific formatting.
        * 2. Sends the payload to the specified Teams webhook URL.
        * 3. Returns the parsed response object.
        *
        * Response Structure:
        * {
        *   "replyId"  : "<<replyId>>",  // Message ID of the reply posted 
                                            (the property will be absent in case of the flow of root post creation).
        *   "threadId" : "<<threadId>>"  // Message ID of the root post (thread ID).
        * }
        */
    static def teamsSend(String teamsWebhookUrl, 
                        String type,
                        String teamsTeamName,
                        String teamsChannelName, 
                        String threadId, 
                        String replyId, 
                        String status, 
                        String msgTitle,
                        String msgBody
 ) {
            //  Step 1: Get the URL
            String url = teamsWebhookUrl

            //  Step 2: Build the Payload for creating the msteams post with Adaptive Card based on the type of message
            def payload = buildTeamsMessagePayloadWithAdaptiveCard(type,
                                                        teamsTeamName,
                                                        teamsChannelName, 
                                                        threadId, 
                                                        replyId, 
                                                        status, 
                                                        msgTitle,
                                                        msgBody)

            // //  Step 3: Call API and return parsed response
             return sendMessageToTeamsUsingWebhook(url, payload)

            
}


   static def teamsSend_authenticated(String teamsWebhookUrl, 
                        String type,
                        String teamsTeamName,
                        String teamsChannelName, 
                        String threadId, 
                        String replyId, 
                        String status, 
                        String msgTitle,
                        String msgBody
 ) {
            //  Step 1: Get the URL
            String url = teamsWebhookUrl

            //  Step 2: Build the Payload for creating the msteams post with Adaptive Card based on the type of message
            def payload = buildTeamsMessagePayloadWithAdaptiveCard(type,
                                                        teamsTeamName,
                                                        teamsChannelName, 
                                                        threadId, 
                                                        replyId, 
                                                        status, 
                                                        msgTitle,
                                                        msgBody)

            // //  Step 3: Call API and return parsed response
            // return sendMessageToTeamsUsingWebhook(url, payload)

            //  Step 3: Call API and return parsed response
            return sendMessageToTeamsUsingWebhook_authenticated(url, payload)
}

/**
 * Fetches an OAuth 2.0 access token using the Client Credentials grant type.
 *
 * This method sends a POST request to the specified token endpoint with the 
 * required client credentials and scope. It is typically used to authenticate 
 * secure API calls where no user interaction is involved (machine-to-machine communication).
 *
 * Parameters:
 * - tokenUrl     : (Required) OAuth token endpoint URL.
 * - clientId     : (Required) Client ID provided by the authorization server.
 * - clientSecret : (Required) Client Secret provided by the authorization server.
 * - scope        : (Required) Scope for which the token is being requested.
 *
 * Workflow:
 * 1. Builds a URL-encoded request body with client credentials.
 * 2. Sends an HTTP POST request to the token endpoint.
 * 3. Parses the JSON response and returns the access token.
 * 4. If the request fails, logs the error and returns null.
 *
 * Returns:
 * - Access token string (if successful).
 * - Null (if token retrieval fails).
 */
static String fetchAccessToken(String tokenUrl,  String clientId, String clientSecret, String scope) {
    
    def url = new URL(tokenUrl)
    def connection = url.openConnection()
    connection.setRequestMethod("POST")
    connection.setDoOutput(true)
    connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded")

    def body = [
        client_id    : clientId,
        client_secret: clientSecret,
        scope        : scope,
        grant_type   : "client_credentials"
    ].collect { k, v -> "${URLEncoder.encode(k, "UTF-8")}=${URLEncoder.encode(v, "UTF-8")}" }
     .join('&')

        // println "fetchAccessToken"
        // println body
    connection.outputStream.withWriter("UTF-8") { writer ->
        writer << body
    }

    def responseCode = connection.responseCode
    def responseStream = (responseCode >= 200 && responseCode < 300) ? connection.inputStream : connection.errorStream
    def responseText = responseStream.text

    if (responseCode == 200) {
        def json = new JsonSlurper().parseText(responseText)
        println json.access_token
        return json.access_token
    } else {
        println "❌ Failed to fetch token. Status: ${responseCode}, Response: ${responseText}"
        return null
    }
}

   
        
  
    /**
    * Sends a POST request to a Microsoft Teams Webhook (or any REST API) with the given JSON payload.
    *
    * @param url     The webhook URL or endpoint to which the POST request is to be sent.
    * @param payload A Map representing the JSON payload that will be sent in the body of the request.
    * @return        A Map representing the parsed JSON response from the server.
    *
    * Usage:
    * - Builds the payload as a JSON string using JsonBuilder.
    * - Sets required headers (Content-Type: application/json).
    * - Sends the POST request and captures the response.
    * - Parses the response text into a Map using JsonSlurper and returns it.
    * - If an error occurs, logs the exception and returns an empty map.
    *
    * Note:
    * - Response code 200/202 is considered successful, and the corresponding response is parsed.
    * - In case of failure, error stream is read and parsed (if possible).
    */

    static Map sendMessageToTeamsUsingWebhook(String url, Map payload) {
        String jsonPayload = new JsonBuilder(payload).toPrettyString()
        def jsonResponse = [:]

        try {
            def connection = new URL(url).openConnection()
            // println url
            // print jsonPayload
            connection.setRequestMethod("POST")
            connection.setRequestProperty("Content-Type", "application/json")
            connection.setDoOutput(true)

            connection.outputStream.withWriter("UTF-8") { writer ->
                writer.write(jsonPayload)
            }

            def responseStream = (connection.responseCode in [200, 202]) ?
                    connection.inputStream : connection.errorStream

            def responseText = responseStream.text
            println "Raw Response: $responseText"

            jsonResponse = new JsonSlurper().parseText(responseText)

        } catch (Exception e) {
            println "Exception during API call: ${e.message}"
        }

        return jsonResponse
    
    }



    /**
 * Builds the final message payload to be sent to Microsoft Teams via webhook.
 * This wraps the Adaptive Card content along with team/channel/thread metadata.
 * Returns a structured payload ready to be sent to Teams via webhook.
 * @param type            Message type: 'post' or 'reply' (controls styling and placement).
    * @param teamsTeamName   Target Teams team name.
    * @param teamsChannelName Target Teams channel name.
    * @param threadId        (Optional) Thread ID for root post.
    * @param replyId         (Optional) Reply ID for nested replies.
    * @param status          Status indicator: 'loading', 'success', or 'failure' (controls icon).
    * @param msgTitle        Title text to display in the card.
    * @param msgBody         Detailed message body.
    * @return                A Map representing the Teams-compatible Adaptive Card payload.
    *
 */
static Map buildTeamsMessagePayloadWithAdaptiveCard(String type,
                                                    String teamsTeamName,
                                                    String teamsChannelName,
                                                    String threadId,
                                                    String replyId,
                                                    String status,
                                                    String msgTitle,
                                                    String msgBody) {

    def adaptiveCardAttachment = constructAdaptiveCardPayloadForTeamsPost(
                                    type, status, msgTitle, msgBody)

    return [
        teamsChannelName: teamsChannelName,
        teamsTeamName   : teamsTeamName,
        type            : type,
        threadId        : threadId,
        replyId         : replyId,
        attachments     : [ adaptiveCardAttachment ]
    ]
}
    
    /**
    * Constructs a Microsoft Teams Adaptive Card payload for either a new post or a reply message.
    *
    * Notes:
    * - Dynamically sets the icon and card styling based on status and message type.
    
    */

static Map constructAdaptiveCardPayloadForTeamsPost(String type,
                                                    String status,
                                                    String msgTitle,
                                                    String msgBody) {

    String iconurl = "https://i.gifer.com/ZKZg.gif"
    String containerStyle = ""
    String textSize = ""
    String fontColor = ""

    // Set icon URL based on status
    switch (status) {
        case 'loading':
            iconurl = "https://i.gifer.com/ZKZg.gif"
            break
        case 'success':
            iconurl = "https://cdn-icons-png.flaticon.com/512/845/845646.png"
            break
        case 'failure':
            iconurl = "https://cdn-icons-png.flaticon.com/512/1828/1828665.png"
            break
    }

    // Set styling based on post type
    switch (type) {
        case 'post':
            containerStyle = "default"
            textSize = "large"
            fontColor = "good"
            break
        case 'reply':
            containerStyle = "accent"
            textSize = "medium"
            fontColor = "default"
            break
    }

    return [
        contentType: "application/vnd.microsoft.card.adaptive",
        content: [
            type    : "AdaptiveCard",
            $schema : "http://adaptivecards.io/schemas/adaptive-card.json",
            version : "1.4",
            msTeams : [ width: "Full" ],
            body    : [
                [
                    type : "Container",
                    style: containerStyle,
                    bleed: true,
                    items: [
                        [
                            type   : "ColumnSet",
                            columns: [
                                [
                                    type : "Column",
                                    width: "auto",
                                    items: [
                                        [
                                            type : "Image",
                                            url  : iconurl,
                                            size : "medium",
                                            width: "20px",
                                            style: "default"
                                        ]
                                    ]
                                ],
                                [
                                    type : "Column",
                                    width: "stretch",
                                    items: [
                                        [
                                            type : "TextBlock",
                                            text : msgTitle,
                                            weight: "bolder",
                                            size : textSize,
                                            wrap : true,
                                            color: fontColor
                                        ]
                                    ]
                                ]
                            ]
                        ]
                    ]
                ],
                [
                    type : "Container",
                    items: [
                        [
                            type   : "ColumnSet",
                            columns: [
                                [
                                    type : "Column",
                                    items: [
                                        [
                                            type : "TextBlock",
                                            text : msgBody,
                                            weight: "bolder",
                                            size : "default",
                                            color: "default",
                                            wrap : true
                                        ]
                                    ]
                                ]
                            ]
                        ]
                    ]
                ]
            ]
        ]
    ]
}

/**
     * Sends a POST request to a Microsoft Teams Webhook (or any secured REST API) with Bearer token authentication.
    *
    * This method sends a JSON-formatted payload to a secured Microsoft Teams webhook 
    * using an access token acquired via OAuth 2.0 client credentials flow.
    *
    * @param url     The webhook URL or endpoint to which the POST request is to be sent.
    * @param payload A Map representing the JSON payload that will be sent in the body of the request.
    * @return        A Map representing the parsed JSON response from the server.
    *
    * Workflow:
    * 1. Builds the JSON string from the provided payload using JsonBuilder.
    * 2. Retrieves an OAuth access token by calling fetchAccessToken().
    * 3. Sets required HTTP headers including Authorization: Bearer token.
    * 4. Sends the POST request with payload and captures the server response.
    * 5. Parses the response into a Map and returns it.
    *
    * Notes:
    * - Successful responses (HTTP 200 or 202) are parsed and returned.
    * - In case of failure or exception, logs the error and returns an empty map.
    */

    static Map sendMessageToTeamsUsingWebhook_authenticated(String url, Map payload) {
        String jsonPayload = new JsonBuilder(payload).toPrettyString()
        def jsonResponse = [:]

        try {


            String accessToken = fetchAccessToken(TOKEN_URL,CLIENT_ID,CLIENT_SECRET,SCOPE)
            println "accessToken"
            println accessToken

            def connection = new URL(url).openConnection()
            // println url
            // print jsonPayload
            connection.setRequestMethod("POST")
            connection.setRequestProperty("Authorization", "Bearer ${accessToken}")
            connection.setRequestProperty("Content-Type", "application/json")
            connection.setDoOutput(true)

            connection.outputStream.withWriter("UTF-8") { writer ->
                writer.write(jsonPayload)
            }

            def responseStream = (connection.responseCode in [200, 202]) ?
                    connection.inputStream : connection.errorStream

            def responseText = responseStream.text
            println "Raw Response: $responseText"

            jsonResponse = new JsonSlurper().parseText(responseText)

        } catch (Exception e) {
            println "Exception during API call: ${e.message}"
        }

        return jsonResponse
    
    }
}