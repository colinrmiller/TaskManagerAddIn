using Newtonsoft.Json;
using System.Collections.Generic;
using static TaskManager.Hevy;

namespace TaskManager.Data
{
    internal class WorkoutResponse
    {
        [JsonProperty("workouts")]
        public List<Workout> Workouts { get; set; }
    }
}
