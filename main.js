import { Scheduler } from './node_modules/@bryntum/scheduler/scheduler.module.js';
import "@bryntum/scheduler/scheduler.stockholm.css";

// get the current date
var today = new Date();

// get the day of the week
var day = today.getDay();

// get the date of the previous sunday
var previousSunday = new Date(today);
previousSunday.setDate(today.getDate() - day);

// get the date of the following Saturday
var nextSaturday = new Date(today);
nextSaturday.setDate(today.getDate() + (6 - day));

const scheduler = new Scheduler({
    appendTo : "scheduler",

    resourceImagePath : 'images/users/',

    eventStyle: 'colored',

    showEventCount : true,

    resourceMargin: 0,

    startDate : previousSunday,
    endDate   : nextSaturday,
    viewPreset : 'dayAndWeek',

    listeners : {
        dataChange: function (event) {
            updateMicrosoft(event);
          }},

    columns : [
        { type : 'resourceInfo', text : 'Name', field : 'name', width : 160 }
    ],

    features : {
        group : 'hasEvent',
        eventEdit : {
            items : {
                // Key to use as fields ref (for easier retrieval later)
                color : {
                    type  : 'combo',
                    label : 'Color',
                    items : ['gray', 'blue', 'purple', 'green', 'pink', 'yellow'],
                    // name will be used to link to a field in the event record when loading and saving in the editor
                    name  : 'eventColor'
                },
                icon : {
                    type  : 'combo',
                    label : 'Icon',

                    items: [
                        { text: 'None', value: '' },
                        { text: 'Door', value: 'b-fa b-fa-door-open' },
                        { text: 'Car', value: 'b-fa b-fa-car' },
                        { text: 'Coffee', value: 'b-fa b-fa-coffee' },
                        { text: 'Envelope', value: 'b-fa b-fa-envelope' },
                    ],

                    name  : 'iconCls' 
                }
            }
        },
    },

    // Custom event renderer, simple version
    eventRenderer({
            eventRecord
        }) {
            if (eventRecord.name == "Open") {
                eventRecord.iconCls = "b-fa b-fa-door-open";
                return `${eventRecord.name}`;
            } else if (eventRecord.name == "Vacation"){
                eventRecord.iconCls = "b-fa b-fa-sun";
                return `${eventRecord.name}`;
            } else if (eventRecord.name == "Second shift"){
                eventRecord.iconCls = "b-fa b-fa-moon";
                return `${eventRecord.name}`;
            } else {
                return `${eventRecord.name}`;
            }
        }
});


async function updateMicrosoft(event) {
    if (event.action == "update") {
        if ("name" in event.changes || "startDate" in event.changes || "endDate" in event.changes || "resourceId" in event.changes || "eventColor" in event.changes) {
            if ("resourceId" in event.changes){
                if (!("oldValue" in event.changes.resourceId)){
                    return;
                }
            } 
            if (Object.keys(event.record.data).indexOf("shiftId") == -1 && Object.keys(event.changes).indexOf("name") !== -1){
                var newShift = createShift(event.record.name, event.record.startDate, event.record.endDate, event.record.resourceId, event.record.eventColor);
                newShift.then(value => {
                    event.record.data["shiftId"] = value.id;
                  });
                scheduler.resourceStore.forEach((resource) => {
                    if (resource.id == event.record.resourceId) {
                        resource.hasEvent = "Assigned";
                    }
            });
            } else {
                if (Object.keys(event.changes).indexOf("resource") !== -1){
                    return;
                }
                updateShift(event.record.shiftId, event.record.resourceId, event.record.name, event.record.startDate, event.record.endDate, event.record.eventColor);
            }
        }
    } else if (event.action == "remove" && "name" in event.records[0].data){
        deleteShift(event.records[0].data.shiftId);
    }
}


const signInButton = document.getElementById("signin");

async function displayUI() {
    await signIn();
  
    // Hide login button and initial UI
    var signInButton = document.getElementById("signin");
    signInButton.style = "display: none";
    var content = document.getElementById("content");
    content.style = "display: block";
  
    var events = await getAllShifts();
    var members = await getMembers();
    members.value.forEach((member) => {
        var user = {id: member.userId, name: member.displayName, hasEvent : "Unassigned"};
        // append user to resources list
        scheduler.resourceStore.add(user);
    });
    events.value.forEach((event) => {
        var shift = {resourceId: event.userId, name: event.sharedShift.displayName, startDate: event.sharedShift.startDateTime, endDate: event.sharedShift.endDateTime, eventColor: event.sharedShift.theme, shiftId: event.id, iconCls: ""};
        
        scheduler.resourceStore.forEach((resource) => {
            if (resource.id == event.userId) {
                resource.hasEvent = "Assigned";
                resource.shiftId = event.id;
            }
        });
        
        // append shift to events list
        scheduler.eventStore.add(shift);
    });
  }

signInButton.addEventListener("click", displayUI);
export { displayUI };