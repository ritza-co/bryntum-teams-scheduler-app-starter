import { Scheduler } from '@bryntum/scheduler/scheduler.module.js';
import { createShift, deleteShift, getAllShifts, getMembers, updateShift } from './graph';
import { signIn } from './auth';

const signInButton = document.getElementById('signin');

// get the current date
const today = new Date();

// get the day of the week
const day = today.getDay();

// get the date of the previous Sunday
const previousSunday = new Date(today);
previousSunday.setDate(today.getDate() - day);

// get the date of the following Saturday
const nextSaturday = new Date(today);
nextSaturday.setDate(today.getDate() + (6 - day));

const scheduler = new Scheduler({
    appendTo    : 'scheduler',
    startDate   : previousSunday,
    endDate     : nextSaturday,
    viewPreset  : 'dayAndWeek',
    workingTime : {
        fromHour : 9,
        toHour   : 17
    },
    listeners : {
        dataChange : function(event) {
            updateMicrosoft(event);
        } },
    columns : [
        { text : 'Name', field : 'name', width : 160 }
    ]
});

async function updateMicrosoft(event) {
    if (event.action == 'update') {
        if ('name' in event.changes || 'startDate' in event.changes || 'endDate' in event.changes || 'resourceId' in event.changes || 'eventColor' in event.changes) {
            if ('resourceId' in event.changes){
                if (!('oldValue' in event.changes.resourceId)){
                    return;
                }
            }
            if (Object.keys(event.record.data).indexOf('shiftId') == -1 && Object.keys(event.changes).indexOf('name') !== -1){
                const newShift = createShift(event.record.name, event.record.startDate, event.record.endDate, event.record.resourceId, event.record.eventColor);
                newShift.then(value => {
                    event.record.data['shiftId'] = value.id;
                });
                scheduler.resourceStore.forEach((resource) => {
                    if (resource.id == event.record.resourceId) {
                        resource.hasEvent = 'Assigned';
                    }
                });
            }
            else {
                if (Object.keys(event.changes).indexOf('resource') !== -1){
                    return;
                }
                updateShift(event.record.shiftId, event.record.resourceId, event.record.name, event.record.startDate, event.record.endDate, event.record.eventColor);
            }
        }
    }
    else if (event.action == 'remove' && 'name' in event.records[0].data){
        deleteShift(event.records[0].data.shiftId);
    }
}

async function displayUI() {
    await signIn();

    // Hide login button and initial UI
    const signInButton = document.getElementById('signin');
    signInButton.style = 'display: none';
    const content = document.getElementById('content');
    content.style = 'display: block';

    const events = await getAllShifts();
    const members = await getMembers();

    // Prepare resources array
    const resources = members.value.map((member) => ({
        id       : member.userId,
        name     : member.displayName,
        hasEvent : 'Unassigned'
    }));

    // Prepare shifts array and update resources
    const shifts = events.value.map((event) => {
        // Update corresponding resource's hasEvent status
        const resource = resources.find(r => r.id === event.userId);
        if (resource) {
            resource.hasEvent = 'Assigned';
            resource.shiftId = event.id;
        }

        return {
            resourceId : event.userId,
            name       : event.sharedShift.displayName,
            startDate  : event.sharedShift.startDateTime,
            endDate    : event.sharedShift.endDateTime,
            eventColor : event.sharedShift.theme,
            shiftId    : event.id,
            iconCls    : ''
        };
    });

    // Load all data at once using loadData (this happens before rendering animations)
    scheduler.resourceStore.data = resources;
    scheduler.eventStore.data = shifts;
}

signInButton.addEventListener('click', displayUI);

export { displayUI };