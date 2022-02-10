// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import {
    Activity,
    ChannelAccount,
    CloudAdapter,
    ConversationReference,
    ConversationParameters,
    MessageFactory,
    TeamsActivityHandler,
    teamsGetChannelId,
    TurnContext,
} from 'botbuilder';

var i = 0;
// var savedConnectorClient
var savedServiceUrl


export class TeamsStartNewThreadInChannel extends TeamsActivityHandler {
    constructor() {
        super();

        this.onMembersAdded( async ( context: TurnContext, next ): Promise<void> => {
            console.log('onMembersAdded')
            // const teamsChannelId = teamsGetChannelId( context.activity );
            const teamsChannelId = '19:1e9f7210f28f43ec8faf1ff53b2f54e2@thread.tacv2'
            const channelAccount = context.activity.from as ChannelAccount;
            const message = MessageFactory.text( 'This will be the first message in a new thread' );
            const newConversation = await this.teamsCreateConversation(context, channelAccount, teamsChannelId, message);

            const botAdapter = context.adapter as CloudAdapter;
            const connectorFactory = context.turnState.get(botAdapter.ConnectorFactoryKey);
            // savedConnectorClient = await connectorFactory.create(context.activity.serviceUrl);
            savedServiceUrl = context.activity.serviceUrl;

            setInterval( async () => {
                try {

                    console.log('Sending notification')
                    const conversationParameters = {
                        bot: channelAccount,
                        channelData: {
                            channel: {
                                id: teamsChannelId
                            }
                        },
                        isGroup: true,
                        activity: MessageFactory.text(`notification ${i++}`),
                    } as ConversationParameters;

                    const connectorClient = await connectorFactory.create(savedServiceUrl);
                    const conversationResourceResponse = await connectorClient.conversations.createConversation( conversationParameters );

                } catch(err) {
                    console.log('Error doing it: ', err);
                }

            }, 30000)

            // await context.adapter.continueConversationAsync(
            //     process.env.MicrosoftAppId,
            //     newConversation[ 0 ],
            //     async ( t ) => {
            //         // await t.sendActivity( MessageFactory.text( 'This will be the first response to the new thread' ) );
            //         await t.sendActivity( 'This will be the first response to the new thread');
            //     });

            await next();
        });

        this.onMessage( async ( context: TurnContext, next ): Promise<void> => {
            console.log('onMessage')
            const teamsChannelId = teamsGetChannelId( context.activity );
            const channelAccount = context.activity.from as ChannelAccount;
            const message = MessageFactory.text( 'This will be the first message in a new thread' );
            const newConversation = await this.teamsCreateConversation( context, channelAccount, teamsChannelId, message );

            await context.adapter.continueConversationAsync(
                process.env.MicrosoftAppId,
                newConversation[ 0 ],
                async ( t ) => {
                    await t.sendActivity( MessageFactory.text( 'This will be the first response to the new thread' ) );
                });

            await next();
        });
    }

    // public async teamsCreateConversation( context: TurnContext, channelAccount: ChannelAccount, teamsChannelId: string, message: Partial<Activity> ): Promise<any> {
    public async teamsCreateConversation( context: TurnContext, channelAccount: ChannelAccount, teamsChannelId: string, message: Partial<Activity> ): Promise<any> {
        console.log('teamsCreateConversation')
        const conversationParameters = {
            bot: channelAccount,
            channelData: {
                channel: {
                    id: teamsChannelId
                }
            },
            isGroup: true,

            activity: message
        } as ConversationParameters;

        const botAdapter = context.adapter as CloudAdapter;
        const connectorFactory = context.turnState.get(botAdapter.ConnectorFactoryKey);
        const connectorClient = await connectorFactory.create(context.activity.serviceUrl);

        const conversationResourceResponse = await connectorClient.conversations.createConversation( conversationParameters );
        const conversationReference = TurnContext.getConversationReference( context.activity ) as ConversationReference;
        conversationReference.conversation.id = conversationResourceResponse.id;
        return [ conversationReference, conversationResourceResponse.activityId ];
    }
}
