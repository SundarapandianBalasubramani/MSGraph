import { getUserPhoto, getUserPhotoAndPresence } from "@/service/user";
import { IPersonaSharedProps, Persona, PersonaPresence, PersonaSize } from "@fluentui/react";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Person, Presence } from "microsoft-graph";
import { useEffect, useState } from "react";

export type IUser = {
    user: Person,
    authProvider?: AuthCodeMSALBrowserAuthenticationProvider
}

export const User = ({ user, authProvider }: IUser) => {
    const [img, setImage] = useState('')
    const [presence, setPresence] = useState<Presence>();
    useEffect(() => {
        const getPhoto = async () => {
            try {                
                const userResponse = await getUserPhotoAndPresence(authProvider!, user.id!);                      
                if(userResponse[1]){     
                    setPresence(userResponse[1]);
                }   
                try{
                    if(userResponse[0]){
                        const url = window.URL || window.webkitURL;
                        const blobUrl = url.createObjectURL(userResponse[0]);
                        setImage(blobUrl);
                    }   
                }catch(e){

                }
                           
               
            } catch (e) {
                //console.log(e);
            }
        };
        getPhoto();
    }, []);
    const getPresence = () => {
        if(presence){
            switch (presence.availability) {
                case 'Available':
                    return PersonaPresence.online;
                case 'DoNotDisturb':
                    return PersonaPresence.dnd;
                case 'Away':
                    return PersonaPresence.away;
                case 'Busy':
                    return PersonaPresence.busy;
                case 'Offline':
                    return PersonaPresence.offline;
                default:
                    return PersonaPresence.none;
            }
        }
        return PersonaPresence.none;
    }
    const persona: IPersonaSharedProps = {
        imageInitials: user.givenName && user.surname ? user.givenName.charAt(0).toUpperCase() + user.surname.charAt(0).toUpperCase() : 'NA',
        text: user.displayName || '',
        secondaryText: user.jobTitle || '',
        tertiaryText: user.department || '',
        optionalText: '',
        presence: getPresence(),
        imageUrl: ''
    };

    return <div style={{ width: 300 }}>
        <Persona
            {...persona}
            size={PersonaSize.size120}
            imageAlt={persona.text}
            imageUrl={img}
        />
    </div>
}