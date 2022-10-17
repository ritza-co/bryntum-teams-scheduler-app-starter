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

    startDate : previousSunday,
    endDate   : nextSaturday,
    viewPreset : 'dayAndWeek',

    features : {
        group : 'hasEvent',
        eventEdit : {
            items : {
                // Key to use as fields ref (for easier retrieval later)
                color : {
                    type  : 'combo',
                    label : 'Color',
                    items : ['red', 'green', 'blue', 'purple', 'indigo', 'orange', 'pink', 'gray', 'black', 'yellow'],
                    // name will be used to link to a field in the event record when loading and saving in the editor
                    name  : 'eventColor'
                }
            }
        },
    },

    listeners : {
        dataChange: function (event) {
            updateMicrosoft(event);
          }},

    columns : [
        { type : 'resourceInfo', text : 'Name', field : 'name', width : 160 },    ]
});


scheduler.showEventCount = true;

async function updateMicrosoft(event) {
    if (event.action == "update") {
        var microsoftShifts = await getAllShifts();

        // check if shift exists in microsoft, if it does, update it, if not, create it
        var eventExists = false;

        if ("name" in event.changes || "startDate" in event.changes || "endDate" in event.changes || "resourceId" in event.changes) {
            for (var i = 0; i < microsoftShifts.value.length; i++) {

                const shift = microsoftShifts.value[i];
                var shiftName = shift.sharedShift.displayName;

                if ("name" in event.changes) {
                    if (event.changes.name.oldValue == shiftName) {
                        eventExists = true;
                        updateShift(shiftId, event.record.resourceId, event.record.name, event.record.startDate, event.record.endDate);
                        return;
                    }
                } else if (event.record.name == shiftName) {
                    eventExists = true;
                    updateShift(shiftId, event.record.resourceId, event.record.name, event.record.startDate, event.record.endDate);
                    return;
                }
            }
        } 
        if (eventExists == false && event.record.originalData.name == "New event") {
            createShift(event.record.name, event.record.startDate, event.record.endDate, event.record.resourceId);
            }
        } else if (event.action == "remove" && "name" in event.records[0].data) {
            const microsoftShifts = await getAllShifts();
            var shiftName = event.records[0].data.name;
            for (var i = 0; i < microsoftShifts.value.length; i++) {
                if (microsoftShifts.value[i].sharedShift.displayName == shiftName) {
                    deleteShift(microsoftShifts.value[i].id);
                    return;
                }
            }
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
    var i = 0;
    members.value.forEach((member) => {
        var user = {id: member.userId, name: member.displayName, hasEvent: "Unassigned"};
        // append user to resources list
        scheduler.resourceStore.add(user);
    });
    events.value.forEach((event) => {
        var shift = {resourceId: event.userId, name: event.sharedShift.displayName, startDate: event.sharedShift.startDateTime, endDate: event.sharedShift.endDateTime, eventColor: event.sharedShift.theme};
        
        // loop through resources to find the correct user id
        scheduler.resourceStore.forEach((resource) => {
            if (resource.id == event.userId) {
                resource.hasEvent = "Assigned";
            }
        });

        // append shift to events list
        scheduler.eventStore.add(shift);
    });
  }

signInButton.addEventListener("click", displayUI);
export { displayUI };
