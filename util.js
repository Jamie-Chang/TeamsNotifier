'use strict';

// Conditions are used to synchronize asynchronous execution.
class Conditions {
    constructor() {
        this.callbacks = {};
    }

    // wait on conditiod <id> to be fulfilled
    wait(id, timeout = null) {
        return new Promise(
            (resolve, reject) => {
                if (timeout !== null) {
                    let wait = setTimeout(() => {
                        delete this.callbacks[id];
                        reject(new Error("Timeout waiting for " + id));
                    }, timeout);
                }
                this.callbacks[id] = resolve;
            }
        );
    }

    // notifiy that condition <id> has been fulfilled
    notify(id, message) {
        return new Promise(
            (resolve, reject) => {
                let callback = this.callbacks[id];
                delete this.callbacks[id];
                if (callback) {
                    callback(message);
                    resolve(message);
                }
                reject(new Error("Nothing waiting for " + id));
            }
        )
    }
}