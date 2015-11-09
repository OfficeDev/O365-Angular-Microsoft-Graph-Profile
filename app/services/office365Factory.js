/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
	angular
		.module('profileApp')
		.factory('office365Factory', office365Factory);

	function office365Factory($log, $http) {
		var office365 = {}; 
 
		// Methods
		office365.getUsers = getUsers;
		office365.getDirectReports = getDirectReports;
		office365.getGroups = getGroups;
		office365.getFiles = getFiles;
		office365.getManager = getManager;
		office365.getProfilePicture = getProfilePicture;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		var baseUrl = 'https://graph.microsoft.com/v1.0/myOrganization';
		
		/**
		 * Gets all users in the tenant.
		 */
		function getUsers() {
			var request = {
				method: 'GET',
				url: baseUrl + '/users'
			};

			return $http(request);
		};
		
		/**
		 * Gets the user's direct reports.
		 */
		function getDirectReports(objectId) {
			var request = {
				method: 'GET',
				url: baseUrl + '/users/' + objectId + '/directReports'
			};

			return $http(request);
		};
		
		/**
		 * Gets the groups the user is a member of.
		 */
		function getGroups(objectId) {
			var request = {
				method: 'GET',
				url: baseUrl + '/users/' + objectId + '/memberOf'
			};

			return $http(request);
		};
		
		/**
		 * Gets the user's files.
		 */
		function getFiles(objectId) {
			var request = {
				method: 'GET',
				url: baseUrl + '/users/' + objectId + '/drive/root/children'
			};

			return $http(request);
		};
		
		/**
		 * Gets the user's manager.
		 */
		function getManager(objectId) {
			var request = {
				method: 'GET',
				url: baseUrl + '/users/' + objectId + '/manager'
			};

			return $http(request);
		};
		
		/**
		 * Gets the user's profile picture.
		 */
		function getProfilePicture(objectId) {
			var request = {
				method: 'GET',
				url: baseUrl + '/users/' + objectId + '/photo/$value'
			};

			return $http(request);
		};

		return office365;
	};
})();

// *********************************************************
//
// O365-Angular-Profile, https://github.com/OfficeDev/O365-Angular-Profile
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************