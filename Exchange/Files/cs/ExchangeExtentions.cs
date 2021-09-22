namespace Terrasoft.Configuration
{

	using System;
	using System.Collections.Generic;
	using Terrasoft.Core;
	using Exchange = Microsoft.Exchange.WebServices.Data;

	#region Class: ExchangeExtentions

	/// <summary>
	/// Class contains EWS api assemply items extention methods.
	/// </summary>
	internal static class ExchangeExtentions
	{

		#region Methods: Public

		/// <summary>
		/// Returns <paramref name="propertyDefinition"/> property value from Exchange storage <paramref name="item"/>.
		/// If property value not reachable, returns default <paramref name="propertyDefinition"/> value.
		/// </summary>
		/// <typeparam name="T">Return value type.</typeparam>
		/// <param name="item"><see cref="Exchange.Item"/> instance.</param>
		/// <param name="propertyDefinition"><see cref="Exchange.PropertyDefinition"/> instance.</param>
		/// <returns><paramref name="propertyDefinition"/> value from <paramref name="item"/> if value reachable,
		/// dafault <paramref name="propertyDefinition"/> value otherwise.</returns>
		public static T SafeGetValue<T>(this Exchange.Item item, Exchange.PropertyDefinition propertyDefinition) {
			T value;
			item.TryGetProperty(propertyDefinition, out value);
			return value;
		}
		
		/// <summary>
		/// Checks if <paramref name="item"/> is empty.
		/// </summary>
		/// <param name="item"><see cref="Exchange.PhysicalAddressEntry"/> instance.</param>
		/// <returns><c>True</c> if <paramref name="item"/> is empty, <c>False</c> otherwise.</returns>
		public static bool IsEmpty(this Exchange.PhysicalAddressEntry item) {
			return item != null && !string.IsNullOrEmpty(item.City + item.CountryOrRegion + item.PostalCode +
				item.State + item.State + item.Street);
		}

		/// <summary>
		/// Tests <see cref="Exchange.ExchangeService"/> connection.
		/// </summary>
		/// <param name="service"><see cref="Exchange.ExchangeService"/> instance.</param>
		/// <param name="userConnection"><see cref="UserConnection"/> instance.</param>
		/// <param name="senderEmailAddress">Sender email address.</param>
		/// <param name="stopOnFirstError">Stop synchronization triggers on first error flag.</param>
		public static void TestConnection(this Exchange.ExchangeService service, UserConnection userConnection = null,
				string senderEmailAddress = "", bool stopOnFirstError = false) {
			SynchronizationErrorHelper helper = SynchronizationErrorHelper.GetInstance(userConnection);
			try {
				var id = new Exchange.FolderId(Exchange.WellKnownFolderName.MsgFolderRoot, senderEmailAddress);
				service.FindFolders(id, new Exchange.FolderView(1));
				helper.CleanUpSynchronizationError(senderEmailAddress);
			} catch (Exception ex) {
				helper.ProcessSynchronizationError(senderEmailAddress, ex, stopOnFirstError);
				throw;
			}
		}

		/// <summary>
		/// In the specified directory <paramref name="folder"/> searches for all subdirectories 
		/// that match the specified filter <paramref name="filter"/> with the corresponding class <paramref name="className"/>
		/// and returns them to the collection <paramref name="list"/>.
		/// </summary>
		/// <param name="list">Filtered Exchange Server catalog collection.</param>
		/// <param name="folder">The directory from which the recursive search is performed.</param>
		/// <param name="className">Class name for search.</param>
		/// <param name="filter">Search filter.</param>
		public static void GetAllFoldersByClass(this List<Exchange.Folder> list, Exchange.Folder folder, string className, Exchange.SearchFilter filter) {
			if (folder.ChildFolderCount > 0 && list != null) {
				Exchange.FindFoldersResults childFolders;
				var folderView = new Exchange.FolderView(folder.ChildFolderCount);
				if (filter == null) {
					childFolders = folder.FindFolders(folderView);
				} else {
					childFolders = folder.FindFolders(filter, folderView);
				}
				foreach (Exchange.Folder childFolder in childFolders) {
					if (childFolder.FolderClass == className) {
						list.Add(childFolder);
					}
					list.GetAllFoldersByClass(childFolder, className, filter);
				}
			}
		}

		/// <summary>
		/// In the specified directory <paramref name="folder"/> searches for all subdirectories 
		/// that match the specified filter <paramref name="filter"/> and returns them to the collection <paramref name="list"/>.
		/// </summary>
		/// <param name="list">Filtered Exchange Server catalog collection.</param>
		/// <param name="folder">The directory from which the recursive search is performed.</param>
		/// <param name="filter">Search filter.</param>
		public static void GetAllFoldersByFilter(this List<Exchange.Folder> list, Exchange.Folder folder, Exchange.SearchFilter filter = null) {
			if (folder.ChildFolderCount > 0 && list != null) {
				Exchange.FindFoldersResults searchedFolders;
				var folderView = new Exchange.FolderView(folder.ChildFolderCount);
				if (filter != null) {
					searchedFolders = folder.FindFolders(filter, folderView);
				} else {
					searchedFolders = folder.FindFolders(folderView);
				}
				list.AddRange(searchedFolders);
				foreach (Exchange.Folder findFoldersResult in searchedFolders) {
					list.GetAllFoldersByFilter(findFoldersResult, filter);
				}
			}
		}

		/// <summary>
		/// Gets the unique identifier of state activity.
		/// </summary>
		/// <param name="taskStatus">Task status.</param>
		/// <returns>The unique identifier of the activity status.</returns>
		public static Guid GetActivityStatus(this Exchange.TaskStatus taskStatus) {
			switch (taskStatus) {
				case Exchange.TaskStatus.InProgress:
					return ActivityConsts.InProgressUId;
				case Exchange.TaskStatus.Completed:
					return ActivityConsts.CompletedStatusUId;
				default:
					return ActivityConsts.NewStatusUId;
			}
		}

		/// <summary>
		/// In the specified directory <paramref name="folder"/>, finds all items <see cref="Exchange.Item"/>,
		/// that match the filter <paramref name="filter"/>, and returns them collection. In the object 
		/// <paramref name="view"/> fixed number of read items in the catalog.
		/// </summary>
		/// <param name="folder">Exchange folder.</param>
		/// <param name="filter">Search filter.</param>
		/// <param name="view">Settings view in the directory search operation.</param>
		/// <returns>Items collection<see cref="Exchange.Item"/>.
		/// </returns>
		public static Exchange.FindItemsResults<Exchange.Item> ReadItems(this Exchange.Folder folder,
			Exchange.SearchFilter filter, Exchange.ItemView view) {
			Exchange.FindItemsResults<Exchange.Item> result =
				filter == null ? folder.FindItems(view) : folder.FindItems(filter, view);
			view.Offset = result.NextPageOffset ?? 0;
			return result;
		}

		#endregion

	}

	#endregion

}
