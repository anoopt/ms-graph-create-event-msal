import * as core from '@actions/core';
import { Event } from "@microsoft/microsoft-graph-types";
import { addBusinessDays, format } from 'date-fns';
import Graph from './graph';

async function run() {

    const clientId: string = core.getInput('clientId', { required: true });
    const clientSecret: string = core.getInput('clientSecret', { required: true });
    const tenantId: string = core.getInput('tenantId', { required: true });
    const start: string = core.getInput('start');
    const end: string = core.getInput('end');
    const subject: string = core.getInput('subject', { required: true });
    const body: string = core.getInput('body', { required: true });
    const userEmail: string = core.getInput('userEmail', { required: true });

    const graph = new Graph(
        clientId,
        clientSecret,
        tenantId
    );

    const nextDay: string = format(addBusinessDays(new Date(), 1), 'yyyy-MM-dd');
    const startTime: string = start ? start : `${nextDay}T12:00:00`;
    const endTime: string = end ? end : `${nextDay}T13:00:00`;

    const event: Event = {
        subject,
        body: {
            contentType: "html",
            content: `${body}<br/>Request submitted around ${format(new Date(), 'dd-MMM-yyyy HH:mm')}`
        },
        start: {
            dateTime: startTime,
            timeZone: "GMT Standard Time"
        },
        end: {
            dateTime: endTime,
            timeZone: "GMT Standard Time"
        }
    };

    const result: any = await graph.createEvent(event, userEmail);
    core.setOutput('event', result);
}

run();