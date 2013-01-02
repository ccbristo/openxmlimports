using System.Collections.Generic;

namespace OpenXmlImports.Tests
{
    public class PropertyRateSet
    {
        public List<StateGroupAssignment> StateGroups { get; private set; }
        public List<SomeStuffWithLimits> Limits { get; private set; }

        public PropertyRateSet()
        {
            this.StateGroups = new List<StateGroupAssignment>();
            this.Limits = new List<SomeStuffWithLimits>();
        }
    }

    public class StateGroupAssignment
    {
        public string State { get; set; }
        public string CompanyEnumKey { get; set; }
        public int StateGroup { get; set; }
    }

    public class SomeStuffWithLimits
    {
        public string ClassCode { get; set; }
        public int Occurrence { get; set; }
        public int Aggregate { get; set; }
    }

    public class LimitSet
    {
        public static readonly LimitSet Limit100200 = new LimitSet(100, 200);
        public static readonly LimitSet Limit300300 = new LimitSet(300, 300);
        public static readonly LimitSet Limit500500 = new LimitSet(500, 500);
        public static readonly LimitSet Limit300600 = new LimitSet(300, 600);
        public static readonly LimitSet Limit5001000 = new LimitSet(500, 1000);
        public static readonly LimitSet Limit10001000 = new LimitSet(1000, 1000);
        public static readonly LimitSet Limit20002000 = new LimitSet(2000, 2000);
        public static readonly LimitSet Limit10002000 = new LimitSet(1000, 2000);
        public static readonly LimitSet Limit20004000 = new LimitSet(2000, 4000);

        public int Occurrence { get; private set; }
        public int Aggregate { get; private set; }

        public LimitSet(int occurrence, int aggregate)
        {
            this.Occurrence = occurrence;
            this.Aggregate = aggregate;
        }

        public static IEnumerable<LimitSet> All
        {
            get
            {
                yield return Limit100200;
                yield return Limit300300;
                yield return Limit500500;
                yield return Limit300600;
                yield return Limit5001000;
                yield return Limit10001000;
                yield return Limit20002000;
                yield return Limit10002000;
                yield return Limit20004000;
            }
        }

        public override int GetHashCode()
        {
            return this.Occurrence;
        }

        public override bool Equals(object obj)
        {
            LimitSet other = obj as LimitSet;

            if (other == null)
                return false;

            return this.Occurrence == other.Occurrence &&
                this.Aggregate == other.Aggregate;
        }
    }
}