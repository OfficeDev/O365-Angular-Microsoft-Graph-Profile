/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
	angular
		.module('profileApp')
		.controller('MainController', MainController);

	MainController.$inject = ['$scope', '$log', '$q', 'adalAuthenticationService', 'office365Factory'];
	
	/**
	 * The MainController code.
	 */
	function MainController($scope, $log, $q, adalAuthenticationService, office365) {
		var vm = this;
		
		// Properties
		vm.users = [];
		vm.activeUser;
		 
		// Methods
		vm.selectUser = selectUser;
		vm.setUserClass = setUserClass;
		vm.setManager = setManager;
		
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		/**
		 * This function does any initialization work the 
		 * controller needs.
		 */
		(function activate() {
			$log.info('MainController invoked.');
			
			// Once the user is logged in, fetch the data.
			if (adalAuthenticationService.userInfo.isAuthenticated) {
				getUsersAsync()
					.then(function () {
						getUserInfo();
					}, function (err) {
						$log.error(err);
					});
			}
		})();
		
		/**
		 * Clears the last selected user's data from the view model
		 * and gets the data for the new selected user.
		 */
		function selectUser() {
			// Switch to first tab.
			vm.makeUserDetailsTabActive = true;
			
			// Clear current data.
			vm.activeUser.directReports = [];
			vm.activeUser.manager = null;
			vm.activeUser.groups = [];
			vm.activeUser.files = [];
			
			// Get user's data from Office 365.
			getUserInfo();
		};
		
		/**
		 * Sets the user's manager as the selected user if a manager exists.
		 */
		function setManager() {
			if (vm.activeUser.manager) {
				vm.activeUser = vm.activeUser.manager;
				vm.selectUser();
			}
		};
		
		/**
		 * Sets class of selected user.
		 */
		function setUserClass(objectId) {
			if (objectId === vm.activeUser.objectId) {
				return 'active';
			}
			else {
				return '';
			}
		};
		
		/**
		 * Gets all users in the tenant.
		 */
		function getUsersAsync() {
			return $q(function (resolve, reject) {
				office365.getUsers()
					.then(function (res) {
						// Bind data to the view model.
						vm.users = res.data.value;	
					
						// Set the selected user as the first user in the directory.		
						vm.activeUser = vm.users[0];
						vm.activeUser.directReports = [];
						vm.activeUser.manager = null;
						vm.activeUser.groups = [];
						vm.activeUser.files = [];

						resolve();
					}, function (err) {
						reject(err);
					});
			});
		};
		
		/**
		 * Gets active user's direct reports, groups, and files.
		 */
		function getUserInfo() {
			// Get the selected user's profile picture.
			office365.getProfilePicture(vm.activeUser.objectId)
				.then(function (res) {	
					// Convert raw image data to encoded data to display.
					console.log(res);
					var imageUrl = "data:image/*;base64," + res.data;
						
					// Bind data to the view model.
					vm.activeUser.photoUrl = imageUrl;

					$log.debug('Photo data: ', res);
					$log.debug('Photo URL: ', vm.activeUser.photoUrl);
				}, function (err) {
					$log.error('User does not have a thumbnail photo.', err);
					vm.activeUser.photoUrl = 'assets/avatar.png';
				});
				
			// Get the selected user's direct reports.
			office365.getDirectReports(vm.activeUser.objectId)
				.then(function (res) {
					// Bind data to the view model.
					vm.activeUser.directReports = res.data.value;
				}, function (err) {
					$log.error(err);
				});
					
			// Get the selected user's manager.
			office365.getManager(vm.activeUser.objectId)
				.then(function (res) {
					// Bind data to the view model.
					vm.activeUser.manager = res.data;
				}, function (err) {
					if (err.status === 404) {
						$log.error('The selected user does not have a manager.');
					}
					else {
						$log.error(err);
					}
				});
					
			// Get the groups the selected user is a member of.
			office365.getGroups(vm.activeUser.objectId)
				.then(function (res) {
					// Bind data to the view model.
					vm.activeUser.groups = res.data.value;
				}, function (err) {
					$log.error(err);
				});
	
			// Get the selected user's files.
			office365.getFiles(vm.activeUser.objectId)
				.then(function (res) {
					// Bind data to the view model.
					vm.activeUser.files = res.data.value;
				}, function (err) {
					$log.error('Unable to get files.', err);
				});
		};
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
