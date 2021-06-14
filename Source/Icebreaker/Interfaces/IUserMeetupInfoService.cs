// <copyright file="IUserMeetupInfoService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Icebreaker.Interfaces
{
    using Icebreaker.Helpers;
    using Microsoft.Graph;
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Contains methods to for Icebreaker bot matching
    /// </summary>
    public interface IUserMeetupInfoService
    {
        /// <summary>
        /// Generate pairups and send pairup notifications.
        /// </summary>
        /// <returns>The number of pairups that were made</returns>
        Task<IList<UserMatchInfoWithCount>> GetUserMeetupLeaders();
    }
}