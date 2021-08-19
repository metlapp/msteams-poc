const { TeamsInfo } = require("botbuilder");
const axios = require('axios');

//TODO Update these
const authAccount = {
    username: 'admin@admin.com',
    password: 'admin'
}
const apiConfig = {
    baseUrl: 'http://127.0.0.1:8000',
    auth: authAccount,
    headers: {
        "Content-type": "application/json"
    },
}

module.exports = class TeamsUtils {

    //Register a team as an Organization with the metl api
	static async registerTeam(context) {
        //Fetch team data
		let {teamData, teamMembers, teamChannels} = await this.getTeamData(context);

        //Ensure the team is registered with the metl api
		let isRegistered = await this.isTeamRegistered(teamData.id);
		if (isRegistered) {
			return await this.fullTeamsUpdate(context);
		}

        //Prep channel and member data
		let channelsAsString = this.channelsToString(teamChannels);
		let membersAsString = this.usersToString(teamMembers);

		//Prep data object for AXIOS call
		var data = {
			"name": teamData.name,
			"users": [],
			"categories": [],
			"slack_installation_object": null,
			"teams_id": teamData.id,
			"teams_tenant_id": teamData.tenantId,
			"teams_user_list": membersAsString,
			"teams_user_dump": teamMembers,
			"teams_channel_list": channelsAsString,
			"teams_channel_dump": teamChannels
		};

		//Prepare AXIOS config
		var config = {
			method: 'post',
			url: `${apiConfig.baseUrl}/organizations/`,
			auth: apiConfig.auth,
			headers: apiConfig.headers,
			data : data
		};

		return new Promise((resolve, reject) => {
			axios(config)
			.then(res => {
				if (res.status === 201) {
					resolve([true, "Organization successfully registered with Metl Solutions."]);
				} else {
					resolve([false, res.data]);
				}
			})
			.catch(error => {
				reject([false, error]);
			});
		});
	}

    //Sends an update to the api. This is only tested with patch and is only used for updating organization info.
    static async updateMetlOrganization(teamID, data) {
        var config = {
			method: 'patch',
			url: `${apiConfig.baseUrl}/organizations/${teamID}/`,
			auth: apiConfig.auth,
			headers: apiConfig.headers,
			data : data
		};

        return new Promise((resolve, reject) => {
            axios(config)
            .then((res) => {
                if (res.status === 200 || res.status === 204) {
                    resolve([true, "Organization updated!"]);
                } else {
                    resolve([false, "There was an error updating the organization."]);
                }
            }).catch((err) => {
                console.error(err);
                reject([false, err]);
            });
        });
    }

	static async fullTeamsUpdate(context) {
        //Fetch team data
        let {teamData, teamMembers, teamChannels} = await this.getTeamData(context);

        //Prep channel and member data
		let channelsAsString = this.channelsToString(teamChannels);
		let membersAsString = this.usersToString(teamMembers);

		//Prep data object for AXIOS call
		var data = {
			"disabled": false,
			"slack_installation_object": null,
			"teams_tenant_id": teamData.tenantId,
			"teams_user_list": membersAsString,
			"teams_user_dump": teamMembers,
			"teams_channel_list": channelsAsString,
			"teams_channel_dump": teamChannels
		};

        //Preform the update request
		let [success, response] = await this.updateMetlOrganization(teamData.id, data);
		if (success) {
			return [true, "Organization data synced with Metl Solutions."];
		}

		return [false, response];
	}

    //When the bot leaves the team, disable it in the metl api
	static async deactivateTeam(context) {
        //Grab teamID from the activity since the bot has technically left
		let teamID = context.activity.conversation.id;

		let data = {
            "disabled": true
        }

        //Preform the update request
        let [success, response] = await this.updateMetlOrganization(teamID, data);
		if (success) {
            console.log(`Organization ${teamID} Disabled`);
		} else {
            console.error(response);
        }

		return;
	}

    //Returns whether or not the team is registered as on organization with Metl
	static async isTeamRegistered(teamsID) {
		return new Promise((resolve, reject) => {
			axios.get(`${apiConfig.baseUrl}/organizations/${teamsID}/`, {
				auth: apiConfig.auth
			})
			.then(res => {
				if (res.data.teams_id && res.data.teams_id != '') {
					resolve(true);
				} else {
					resolve(false);
				}
			})
			.catch(error => {
				console.error(error);
				reject(false);
			});
		});
	}

    //Update the channels in the metl api to track new, removed, or edited channels
    static async updateChannels(context) {
        //Fetch team data
		let {teamData, teamChannels} = await this.getTeamData(context);

		//Ensure the team is registered with the metl api
		let isRegistered = await this.isTeamRegistered(teamData.id);
		if (!isRegistered) {
			return await this.registerTeam(context);
		}

		//Convert the channels array to a comma-delimited string of names
		let channelsAsString = this.channelsToString(teamChannels);

		let data = {
            "teams_channel_list": channelsAsString,
            "teams_channel_dump": teamChannels
        };

        //Preform the update request
		let [success, response] = await this.updateMetlOrganization(teamData.id, data);
		if (success) {
			return [true, "Channels updated for the organization."];
		}

		return [false, response];
	}

	//Update the members variable to track new/removed users
	static async updateMembers(context) {
		//Fetch team data
		let {teamData, teamMembers} = await this.getTeamData(context);

		//Ensure the team is registered with the metl api
		let isRegistered = await this.isTeamRegistered(teamData.id);
		if (!isRegistered) {
			return await this.registerTeam(context);
		}

		//Convert the members array to a comma-delimited string of first+last names
		let usersAsString = this.usersToString(teamMembers);

        let data = {
            "teams_user_list": usersAsString,
            "teams_user_dump": teamMembers
        }

		//Preform the update request
		let [success, response] = await this.updateMetlOrganization(teamData.id, data);
		if (success) {
			return [true, "Members updated for the organization."];
		}

		return [false, response];
	}

    //Takes in an array of channel objects and returns a comma-delimited string
	static channelsToString(channels) {
		let channelsStr = "";
		channels.forEach((cur, index) => {
			if (index != 0) {
				channelsStr += ",";
			}

			channelsStr += cur.name;
		});

		return channelsStr;
	}

    //Takes in an array of user objects and returns a comma-delimited string
    static usersToString(users) {
		let usersStr = "";
		users.forEach((cur, index) => {
			if (index != 0) {
				usersStr += ",";
			}

			usersStr += cur.name;
		});

		return usersStr;
	}

    //Returns an object contains the team details, team members, and team channels
    static async getTeamData(context) {
		let teamData = await TeamsInfo.getTeamDetails(context);
		let teamMembers = await TeamsInfo.getTeamMembers(context);
		let teamChannels = await TeamsInfo.getTeamChannels(context);

		//The first channel every team has is a general channel and for some reason it doesn't have a name
		if (!teamChannels[0].name && teamChannels[0].id === teamData.id) {teamChannels[0].name = "General"};

		return {
			teamData: teamData,
			teamMembers: teamMembers,
			teamChannels: teamChannels
		};
	}
};