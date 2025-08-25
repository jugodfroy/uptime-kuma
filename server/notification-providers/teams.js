const NotificationProvider = require("./notification-provider");
const axios = require("axios");
const { setting } = require("../util-server");
const { DOWN, UP, getMonitorRelativeURL } = require("../../src/util");

class Teams extends NotificationProvider {
    name = "teams";

    /**
     * Generate the message to send
     * @param {const} status The status constant
     * @param {string} monitorName Name of monitor
     * @param {boolean} withStatusSymbol If the status should be prepended as symbol
     * @returns {string} Status message
     */
    _statusMessageFactory = (status, monitorName, withStatusSymbol) => {
        if (status === DOWN) {
            return (withStatusSymbol ? "🔴 " : "") + `[${monitorName}] went down`;
        } else if (status === UP) {
            return (withStatusSymbol ? "✅ " : "") + `[${monitorName}] is back online`;
        }
        return "Notification";
    };

    /**
     * Select the style to use based on status
     * @param {const} status The status constant
     * @returns {string} Selected style for adaptive cards
     */
    _getStyle = (status) => {
        if (status === DOWN) {
            return "attention";
        }
        if (status === UP) {
            return "good";
        }
        return "default";
    };

    /**
     * Generate payload for notification
     * @param {object} args Method arguments
     * @param {object} args.heartbeatJSON Heartbeat details
     * @param {string} args.monitorName Name of the monitor affected
     * @param {string} args.monitorUrl URL of the monitor affected
     * @param {string} args.dashboardUrl URL of the dashboard affected
     * @returns {object} Notification payload
     */
    _notificationPayloadFactory = ({
        heartbeatJSON,
        monitorName,
        monitorUrl,
        dashboardUrl,
        is_legacyMsTeamsWebhook,        
    }) => {
        const status = heartbeatJSON?.status;
        const body = [];
        const facts = []; //for legacy url only
        const actions = [];

        body.push({
            type: "TextBlock",
            text: this._statusMessageFactory(status, monitorName, true), 
            weight: "Bolder",
            size: "Large",
            color: this._getStyle(status)
        });

        if (heartbeatJSON?.msg) {
            body.push({
                type: "TextBlock",
                text: `**Description:** ${heartbeatJSON.msg}`,
                wrap: true
            });
        }

        if (monitorName) {
            body.push({
                type: "TextBlock",
                text: `**Monitor:** ${monitorName}`,
                wrap: true
            });
        }

        if (monitorUrl && monitorUrl !== "https://") {
            body.push({
                type: "TextBlock",
                text: `**URL:** [${monitorUrl}](${monitorUrl})`,
                wrap: true
            });
        }

        if (heartbeatJSON?.localDateTime) {
            body.push({
                type: "TextBlock",
                text: `**Time:** ${heartbeatJSON.localDateTime}${heartbeatJSON.timezone ? ` (${heartbeatJSON.timezone})` : ""}`,
                wrap: true
            });
        }

        if (dashboardUrl) {
            actions.push({
                type: "Action.OpenUrl",
                title: "Visit Uptime Kuma",
                url: dashboardUrl
            });
        }

        actions.push({
            type: "Action.OpenUrl",
            title: "Go to Grafana",
            url: "https://027cc856-399c-4091-9da7-7cbecd1206a9.dashboard.obs.fr-par.scw.cloud/d/b281712d-8bff-41ef-9f3f-71ad43c05e9b/vllm-kube?orgId=1&var-DS_PROMETHEUS=ad06aec8-07c6-45d5-9e72-c2d52b4d2bd0&var-model_name=Qwen%2FQwen3-32B-FP8&var-cluster=llm-cluster&from=now-30m&to=now"
        });

        const payload = {
            type: "message",
            attachments: [
                {
                    contentType: "application/vnd.microsoft.card.adaptive",
                    contentUrl: null,
                    content: {
                        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                        type: "AdaptiveCard",
                        version: "1.4",
                        body: body,
                        actions: actions
                    }
                }
            ]
        };

        return payload;
    };

    /**
     * Generate payload for notification
     * @param {object} args Method arguments
     * @param {object} args.heartbeatJSON Heartbeat details
     * @param {string} args.monitorName Name of the monitor affected
     * @param {string} args.monitorUrl URL of the monitor affected
     * @param {string} args.dashboardUrl URL of the dashboard affected
     * @returns {object} Notification payload
     */
    _legacyNotificationPayloadFactory = ({
        heartbeatJSON,
        monitorName,
        monitorUrl,
        dashboardUrl,
    }) => {
        const status = heartbeatJSON?.status;
        const facts = [];
        const actions = [];

        if (dashboardUrl) {
            actions.push({
                "type": "Action.OpenUrl",
                "title": "Visit Uptime Kuma",
                "url": dashboardUrl
            });
        }

        if (heartbeatJSON?.msg) {
            facts.push({
                title: "Description",
                value: heartbeatJSON.msg,
            });
        }

        if (monitorName) {
            facts.push({
                title: "Monitor",
                value: monitorName,
            });
        }

        if (monitorUrl && monitorUrl !== "https://") {
            facts.push({
                title: "URL",
                // format URL as markdown syntax, to be clickable
                value: `[${monitorUrl}](${monitorUrl})`,
            });
            actions.push({
                "type": "Action.OpenUrl",
                "title": "Visit Monitor URL",
                "url": monitorUrl
            });
        }

        if (heartbeatJSON?.localDateTime) {
            facts.push({
                title: "Time",
                value: heartbeatJSON.localDateTime + (heartbeatJSON.timezone ? ` (${heartbeatJSON.timezone})` : ""),
            });
        }

        const payload = {
            "type": "message",
            // message with status prefix as notification text
            "summary": this._statusMessageFactory(status, monitorName, true),
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "contentUrl": "",
                    "content": {
                        "type": "AdaptiveCard",
                        "body": [
                            {
                                "type": "Container",
                                "verticalContentAlignment": "Center",
                                "items": [
                                    {
                                        "type": "ColumnSet",
                                        "style": this._getStyle(status),
                                        "columns": [
                                            {
                                                "type": "Column",
                                                "width": "auto",
                                                "verticalContentAlignment": "Center",
                                                "items": [
                                                    {
                                                        "type": "Image",
                                                        "width": "32px",
                                                        "style": "Person",
                                                        "url": "https://raw.githubusercontent.com/louislam/uptime-kuma/master/public/icon.png",
                                                        "altText": "Uptime Kuma Logo"
                                                    }
                                                ]
                                            },
                                            {
                                                "type": "Column",
                                                "width": "stretch",
                                                "items": [
                                                    {
                                                        "type": "TextBlock",
                                                        "size": "Medium",
                                                        "weight": "Bolder",
                                                        "text": `**${this._statusMessageFactory(status, monitorName, false)}**`,
                                                    },
                                                    {
                                                        "type": "TextBlock",
                                                        "size": "Small",
                                                        "weight": "Default",
                                                        "text": "Uptime Kuma Alert",
                                                        "isSubtle": true,
                                                        "spacing": "None"
                                                    }
                                                ]
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "FactSet",
                                "separator": false,
                                "facts": facts
                            }
                        ],
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "version": "1.5"
                    }
                }
            ]
        };

        if (actions) {
            payload.attachments[0].content.body.push({
                "type": "ActionSet",
                "actions": actions,
            });
        }

        return payload;
    };


    /**
     * Send the notification
     * @param {string} webhookUrl URL to send the request to
     * @param {object} payload Payload generated by _notificationPayloadFactory or _legacyNotificationPayloadFactory
     * @returns {Promise<void>}
     */
    _sendNotification = async (webhookUrl, payload) => {
        await axios.post(webhookUrl, payload);
    };

    /**
     * Send a general notification
     * @param {string} webhookUrl URL to send request to
     * @param {string} msg Message to send
     * @param {boolean} is_legacyMsTeamsWebhook Whether the WebhookUrl is legacy
     * @returns {Promise<void>}
     */
    _handleGeneralNotification = (webhookUrl, msg, is_legacyMsTeamsWebhook) => {
        let payload = {}
        if (is_legacyMsTeamsWebhook){
            payload = this._legacyNotificationPayloadFactory({
                heartbeatJSON: {
                    msg: msg
                }, 
            });
        }
        else {
            payload = this._notificationPayloadFactory({
                heartbeatJSON: {
                    msg: msg
                }, 
            });
        }
        return this._sendNotification(webhookUrl, payload);
    };

    /**
     * Check whether the webhookUrl is a legacy webhook (depreciated)
     * @param {string} webhookUrl URL to send request to
     * @returns {boolean}
     */
    _legacyWebhookUrlChecker = (webhookUrl) => {
        const legacyPattern = /^https:\/\/outlook\.office\.com\/webhook\//; //pattern of legacy MsTeams webhook
        if (legacyPattern.test(webhookUrl)) {
            return true;
        }
        else{
            return false;
        }
    };

    /**
     * @inheritdoc
     */
    async send(notification, msg, monitorJSON = null, heartbeatJSON = null) {
        const okMsg = "Sent Successfully.";
        const is_legacyMsTeamsWebhook = this._legacyWebhookUrlChecker(notification.webhookUrl);
        

        try {
            if (heartbeatJSON == null) {
                await this._handleGeneralNotification(notification.webhookUrl, msg, is_legacyMsTeamsWebhook);
                return okMsg;
            }

            const baseURL = await setting("primaryBaseURL");
            let dashboardUrl;
            if (baseURL) {
                dashboardUrl = baseURL + getMonitorRelativeURL(monitorJSON.id);
            }

            const payload = this._notificationPayloadFactory({
                heartbeatJSON: heartbeatJSON,
                monitorName: monitorJSON.name,
                monitorUrl: this.extractAddress(monitorJSON),
                dashboardUrl: dashboardUrl,
                is_legacyMsTeamsWebhook: is_legacyMsTeamsWebhook
            });

            await this._sendNotification(notification.webhookUrl, payload);
            return okMsg;
        } catch (error) {
            this.throwGeneralAxiosError(error);
        }
    }
}

module.exports = Teams;
