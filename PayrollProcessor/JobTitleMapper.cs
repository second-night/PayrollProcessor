using static PayrollProcessor.Program;

namespace PayrollProcessor
{
    /// <summary>
    /// Maps ADP Workforce Now job title codes to <see cref="Jobs"/> values and back.
    /// DRCOA and DNU are never used for imports.
    /// </summary>
    internal static class JobTitleMapper
    {
        private const float MinimumHoursForJobTitleChange = 50f;

        public const string NonCdlDriver = "NCDL";
        public const string Admin = "ADMIN";
        public const string BodyShop = "BDYSHP";
        public const string DoNotUse = "DNU";
        public const string DriverCharter = "DRCCHSC";
        public const string DriverCoach = "DRCOA";
        public const string DriverDailySchool = "DRDLYSC";
        public const string Mechanic = "MECHNC";
        public const string Para = "PARA";
        public const string WashBay = "WSHBY";

        private static readonly HashSet<string> CodesNeverImported = new(StringComparer.OrdinalIgnoreCase)
        {
            DriverCoach,
            DoNotUse
        };

        private static readonly HashSet<Jobs> JobsExcludedFromMajority = new()
        {
            Jobs.HOLIDAY,
            Jobs.VACATION,
            Jobs.TRAINING,
            Jobs.SALARY
        };

        public static bool IsNeverImported(string? jobTitleCode) =>
            !string.IsNullOrWhiteSpace(jobTitleCode) && CodesNeverImported.Contains(jobTitleCode.Trim());

        public static bool TryMapCodeToJob(string? jobTitleCode, out Jobs job)
        {
            job = default;
            if (string.IsNullOrWhiteSpace(jobTitleCode))
            {
                return false;
            }

            switch (jobTitleCode.Trim().ToUpperInvariant())
            {
                case NonCdlDriver:
                    job = Jobs.NON_CDL_DRIVER;
                    return true;
                case Admin:
                    job = Jobs.ADMIN;
                    return true;
                case BodyShop:
                    job = Jobs.BODY_SHOP;
                    return true;
                case DriverCharter:
                    job = Jobs.DRIVER_LOCAL_SCHOOL_CHARTERS;
                    return true;
                case DriverCoach:
                case DriverDailySchool:
                    job = Jobs.DRIVER_SCHOOL;
                    return true;
                case Mechanic:
                    job = Jobs.MECHANIC;
                    return true;
                case Para:
                    job = Jobs.AIDE_SCHOOL;
                    return true;
                case WashBay:
                    job = Jobs.WASH_BAY;
                    return true;
                case DoNotUse:
                    return false;
                default:
                    return false;
            }
        }

        public static bool TryMapJobToImportCode(Jobs job, out string jobTitleCode)
        {
            switch (CanonicalJobForTitle(job))
            {
                case Jobs.NON_CDL_DRIVER:
                    jobTitleCode = NonCdlDriver;
                    return true;
                case Jobs.ADMIN:
                    jobTitleCode = Admin;
                    return true;
                case Jobs.BODY_SHOP:
                    jobTitleCode = BodyShop;
                    return true;
                case Jobs.DRIVER_SCHOOL:
                    jobTitleCode = DriverDailySchool;
                    return true;
                case Jobs.DRIVER_LOCAL_SCHOOL_CHARTERS:
                    jobTitleCode = DriverCharter;
                    return true;
                case Jobs.MECHANIC:
                    jobTitleCode = Mechanic;
                    return true;
                case Jobs.AIDE_SCHOOL:
                    jobTitleCode = Para;
                    return true;
                case Jobs.WASH_BAY:
                    jobTitleCode = WashBay;
                    return true;
                default:
                    jobTitleCode = "";
                    return false;
            }
        }

        public static Jobs CanonicalJobForTitle(Jobs job) => job switch
        {
            Jobs.DRIVER_COACH => Jobs.DRIVER_SCHOOL,
            Jobs.DRIVER_LOCAL_SCHOOL_CHARTERS
                or Jobs.DRIVER_CHARTER_PRIVATE
                or Jobs.DRIVER_OUT_OF_TOWN_CHARTER => Jobs.DRIVER_LOCAL_SCHOOL_CHARTERS,
            Jobs.AIDE_CHARTER => Jobs.AIDE_SCHOOL,
            Jobs.WASH_BAY_OT => Jobs.WASH_BAY,
            _ => job
        };

        public static bool TryGetMajorityJob(
            Employee employee,
            IReadOnlyDictionary<Jobs, float> historicalHoursByJob,
            out Jobs majorityJob,
            out float majorityHours,
            out float totalJudgedHours)
        {
            majorityJob = default;
            majorityHours = 0f;
            totalJudgedHours = 0f;
            Dictionary<Jobs, float> hoursByJob = new();

            void AddHours(Jobs job, float hours)
            {
                if (JobsExcludedFromMajority.Contains(job) || hours <= 0.01f)
                {
                    return;
                }

                Jobs canonicalJob = CanonicalJobForTitle(job);
                if (JobsExcludedFromMajority.Contains(canonicalJob))
                {
                    return;
                }

                hoursByJob[canonicalJob] = hoursByJob.GetValueOrDefault(canonicalJob) + hours;
            }

            foreach ((Jobs job, float hours) in historicalHoursByJob)
            {
                AddHours(job, hours);
            }

            foreach (Shift shift in employee.Shifts.Where(shift => !shift.IsATotalsShift))
            {
                float hours = shift.JobType == Jobs.DRIVER_COACH
                    ? Math.Max(shift.AllHours(false), shift.CoachTripDays * 8f)
                    : shift.AllHours(false);
                AddHours(shift.JobType, hours);
            }

            if (hoursByJob.Count == 0)
            {
                return false;
            }

            totalJudgedHours = hoursByJob.Values.Sum();
            KeyValuePair<Jobs, float> top = hoursByJob.OrderByDescending(entry => entry.Value).First();
            majorityJob = top.Key;
            majorityHours = top.Value;
            return true;
        }

        public static void EvaluateEmployees(IEnumerable<Employee> employees, EmployeePayrollHistory payrollHistory)
        {
            foreach (Employee employee in employees)
            {
                employee.RecommendedJobTitleCode = "";
                if (employee.IsSalaried || employee.IsTerminated)
                {
                    continue;
                }

                if (!TryGetMajorityJob(employee, payrollHistory. GetHoursByJobFromHistory(employee.IdNumber),
                        out Jobs majorityJob, out float majorityHours, out float totalJudgedHours)
                    || totalJudgedHours < MinimumHoursForJobTitleChange
                    || !TryMapJobToImportCode(majorityJob, out string recommendedCode))
                {
                    continue;
                }

                if (IsNeverImported(recommendedCode))
                {
                    continue;
                }

                string currentCode = (employee.JobTitleCode ?? "").Trim().ToUpperInvariant();
                string recommendedNormalized = recommendedCode.ToUpperInvariant();
                bool currentIsInvalid = string.IsNullOrWhiteSpace(currentCode) || IsNeverImported(currentCode);
                bool currentMatchesRecommended = string.Equals(currentCode, recommendedNormalized, StringComparison.OrdinalIgnoreCase);
                bool currentMatchesMajorityFamily =
                    TryMapCodeToJob(currentCode, out Jobs currentJob)
                    && TryMapJobToImportCode(currentJob, out string currentMappedCode)
                    && string.Equals(currentMappedCode, recommendedNormalized, StringComparison.OrdinalIgnoreCase);

                if (!currentIsInvalid && (currentMatchesRecommended || currentMatchesMajorityFamily))
                {
                    continue;
                }

                employee.RecommendedJobTitleCode = recommendedNormalized;
                Log($"Job title update for {employee.Name} ({employee.IdNumber}): "
                    + $"'{currentCode}' -> '{recommendedNormalized}' based on majority hours "
                    + $"({majorityHours:0.##} of {totalJudgedHours:0.##}) in {majorityJob}.");
            }
        }
    }
}
