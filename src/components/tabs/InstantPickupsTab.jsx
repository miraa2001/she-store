import { useState } from "react";
import { UserPlus } from "lucide-react";
import { INSTANT_PICKUP_LOCATIONS } from "../../lib/instantPickups";
import { PICKUP_POINT } from "../../lib/pickup";
import InstantPickupSection from "../pickup/InstantPickupSection";

export default function InstantPickupsTab({ role, search, onAdd, onDialogChange }) {
  const [location, setLocation] = useState("all");
  const availableLocations = INSTANT_PICKUP_LOCATIONS.filter((item) => role === "rahaf" || item.value !== PICKUP_POINT);
  const visibleLocations = availableLocations.filter((item) => location === "all" || location === item.value);

  return (
    <div className="instant-pickups-tab">
      <div className="instant-pickups-tab-head">
        <h2>استلام فوري</h2>
        <div className="instant-pickups-tab-tools">
          <label>نقطة الاستلام
            <select value={location} onChange={(event) => setLocation(event.target.value)}>
              <option value="all">كل نقاط الاستلام</option>
              {availableLocations.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}
            </select>
          </label>
          {role === "rahaf" ? <button type="button" className="instant-pickup-command" onClick={onAdd}><UserPlus size={18} />اضافة مستلم فوري</button> : null}
        </div>
      </div>
      {visibleLocations.map((item) => (
        <InstantPickupSection key={item.value} pickupPoint={item.value} role={role} title={item.label} search={search} showCreateAction={false} onDialogChange={onDialogChange} />
      ))}
    </div>
  );
}
