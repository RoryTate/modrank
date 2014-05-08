﻿using System.Collections.Generic;

namespace POEApi.Model
{
    internal class GearTypeFactory
    {
        private static List<GearTypeRunner> runners = new List<GearTypeRunner>()
        {
            { new HelmetRunner() },
            { new RingRunner() },
            { new AmuletRunner() },
            { new BeltRunner() },
            { new FlaskRunner() },
            { new MapRunner() },
            { new GloveRunner() },
            { new BootRunner() },
            { new AxeRunner() },
            { new ClawRunner() },
            { new BowRunner() },
            { new DaggerRunner() },
            { new MaceRunner() },
            { new QuiverRunner() },
            { new SceptreRunner() },
            { new StaffRunner() },
            { new SwordRunner() },
            { new ShieldRunner() },
            { new WandRunner() },
            { new ChestRunner() } //Must always be last!
        };

        public static GearType GetType(Gear item)
        {
            foreach (var runner in runners)
                if (runner.IsCompatibleType(item))
                    return runner.Type;

            return GearType.Unknown;
        }

        public static string GetBaseType(Gear item)
        {
            foreach (var runner in runners)
            {
                // If we know the GearType of the item, only query the GearTypeRunner for
                // that type.  If the GearType is unknown, query all of them.
                if (item.GearType == GearType.Unknown || item.GearType == runner.Type)
                {
                    string baseType = runner.GetBaseType(item);
                    if (!string.IsNullOrWhiteSpace(baseType))
                    {
                        return baseType;
                    }
                }
            }
            return null;
        }
    }
}
