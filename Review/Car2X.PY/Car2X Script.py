"""
This Car2X (C2X) example demonstrates how to model communication between vehicles.
At simulation second 200, there is a breakdown of a vehicle. At the time of breakdown,
the vehicle sends out a warning message.
Vehicles receiving this message will drop their speed and adjust their driving behavior
until they passed the incident.
"""

def Initialization():
    # Global Parameters:
    global distDistr
    global Vehicle_Type_C2X_no_message
    global Vehicle_Type_C2X_HasCurrentMessage
    global speed_incident

    distDistr = 1 # number of Distance distribution used for sending out a C2X message
    Vehicle_Type_C2X_no_message = '101' # number of C2X vehicle type (no active message) has to be a string!
    Vehicle_Type_C2X_HasCurrentMessage = '102' # number of C2X vehicle type with active message has to be a string!
    speed_incident = 80 # Speed of vehicles receiving the C2X message in kph


def Main():
    """
    Main control
    """
    # Get several attributes of all vehicles:
    Veh_attributes = Vissim.Net.Vehicles.GetMultipleAttributes(('RoutDecType', 'RoutDecNo', 'VehType', 'No'))

    if Veh_attributes: # Check if there are any vehicles in the network:

        # Filter by VehType C2X:
        Veh_C2X_attributes = [item for item in Veh_attributes if item[2] == Vehicle_Type_C2X_no_message or item[2] == Vehicle_Type_C2X_HasCurrentMessage]

        # For all C2X vehicles: check if there is an incident | incident is modeled as parking routing decision #1
        for cnt_C2X_veh in range(len(Veh_C2X_attributes)):
            if Veh_C2X_attributes[cnt_C2X_veh][0] == 'PARKING' and Veh_C2X_attributes[cnt_C2X_veh][1] == 1: # vehicle has an incident (parking routing decision #1)
                Veh_sending_Msg = Vissim.Net.Vehicles.ItemByKey(Veh_C2X_attributes[cnt_C2X_veh][3])
                Coord_Veh = Veh_sending_Msg.AttValue('CoordFront') # reading the world coordinates (x y z) of the vehicle
                PositionXYZ = Coord_Veh.split(" ")

                Pos_Veh_SM = Veh_sending_Msg.AttValue('Pos') # relative position on the current link
                Veh_sending_Msg.SetAttValue('C2X_HasCurrentMessage', 1)
                Veh_sending_Msg.SetAttValue('C2X_SendingMessage', 1)
                Veh_sending_Msg.SetAttValue('C2X_MessageOrigin', Pos_Veh_SM)

                # Getting vehicles which receive the message:
                Veh_Rec_Message = Vissim.Net.Vehicles.GetByLocation(PositionXYZ[0], PositionXYZ[1], distDistr)

                # Reading Attribute of all Vehicles who are receiving the C2X message (Note: all vehicle classes involved, also non C2X vehicles)
                Attributes = ('Pos', 'VehType', 'C2X_HasCurrentMessage', 'C2X_MessageOrigin', 'C2X_Message', 'DesSpeed', 'C2X_DesSpeedOld')
                Veh_attributes_Rec_Message = list(Veh_Rec_Message.GetMultipleAttributes(Attributes))

                # Adjusting the attributes of the C2X vehicles because of this message:
                for cnt_Veh_Rec_Message in range(len(Veh_attributes_Rec_Message)):
                    atts_current = Veh_attributes_Rec_Message[cnt_Veh_Rec_Message]
                    pos_cur = atts_current[0]
                    veh_type_cur = atts_current[1]
                    pos_C2X_cur = atts_current[3]
                    des_speed_cur = atts_current[5]
                    des_speed_old_cur = atts_current[6]

                    # check if vehicle has C2X & position of C2X message is downstream & there is no other further downsteam message active
                    if veh_type_cur in (Vehicle_Type_C2X_no_message, Vehicle_Type_C2X_HasCurrentMessage) \
                    and pos_cur < Pos_Veh_SM \
                    and (pos_C2X_cur is None or Pos_Veh_SM > pos_C2X_cur):
                        if des_speed_cur == speed_incident:
                            # if the attribute 'DesSpeed' was already set to 'speed_incident', don't overwrite 'C2X_DesSpeedOld' with with current 'DesSpeed' = 'speed_incident'
                            Veh_attributes_Rec_Message[cnt_Veh_Rec_Message] = tuple([int(Vehicle_Type_C2X_HasCurrentMessage), 1, Pos_Veh_SM, 'Breakdown Vehicle ahead!', speed_incident, des_speed_old_cur])
                        else:
                            Veh_attributes_Rec_Message[cnt_Veh_Rec_Message] = tuple([int(Vehicle_Type_C2X_HasCurrentMessage), 1, Pos_Veh_SM, 'Breakdown Vehicle ahead!', speed_incident, des_speed_cur])
                    else:
                        Veh_attributes_Rec_Message[cnt_Veh_Rec_Message] = atts_current[1:] # no changes, vehicle has no C2X or is not affected due to the position
                # Giving back the adjusted attributes to Vissim (note: attribute 'Pos' is read-only)
                Veh_Rec_Message.SetMultipleAttributes(Attributes[1:], Veh_attributes_Rec_Message)

        # Check if vehicles with active message passed the position of the warning message:
        Attributes = ('Pos', 'VehType', 'C2X_HasCurrentMessage', 'C2X_MessageOrigin', 'C2X_Message', 'DesSpeed', 'C2X_DesSpeedOld')
        Veh_attributes = list(Vissim.Net.Vehicles.GetMultipleAttributes(Attributes))

        for cnt_Veh in range(len(Veh_attributes)):
            atts_current = Veh_attributes[cnt_Veh]
            pos_cur = atts_current[0]
            veh_type_cur = atts_current[1]
            C2X_msg_active_cur = atts_current[2]
            pos_C2X_cur = atts_current[3]
            des_speed_old_cur = atts_current[6]
            # if the vehicle has an active C2X message AND the position is larger than the C2X Position
            if C2X_msg_active_cur == 1 and pos_cur > pos_C2X_cur and not (des_speed_old_cur is None) :
                Veh_attributes[cnt_Veh] = [int(Vehicle_Type_C2X_no_message), 0, '', '', des_speed_old_cur, '']
            else:
                Veh_attributes[cnt_Veh] = atts_current[1:] # no changes
    	# Returning the adjusted attributes to Vissim (note: attribute 'Pos' is read-only)
        Vissim.Net.Vehicles.SetMultipleAttributes(Attributes[1:], Veh_attributes)
