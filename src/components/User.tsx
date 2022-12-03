import { getUserPhoto } from "@/service/user";
import { IPersonaSharedProps, Persona, PersonaPresence, PersonaSize } from "@fluentui/react";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Person, Presence } from "microsoft-graph";
import { useEffect, useState } from "react";
import { PersonWithPresence } from "./Directory";

export type IUser = {
    user: PersonWithPresence,
    authProvider?: AuthCodeMSALBrowserAuthenticationProvider
}

export const User = ({ user, authProvider }: IUser) => {
    const [img, setImage] = useState('');
    const getPhoto = async () => {
        try {
            const id = 'userId' in user  ? (user as any).userId : user.id!;
            const photo = await getUserPhoto(authProvider!, user.id!);
            if (photo) {
                const url = window.URL || window.webkitURL;
                const blobUrl = url.createObjectURL(photo);
                setImage(blobUrl);
            }
        } catch (e) {
            //onsole.log(user.givenName, e);
            //Ignore this error normally code will reach here if person photo is not available
        }
    };
    useEffect(() => {
        getPhoto();
    }, []);
    const getPresence = () => {
        if (user.availability) {
            switch (user.availability) {
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
    

    return <div style={{ width: 300 }}>
        <Persona
            imageInitials={user.givenName && user.surname ? user.givenName.charAt(0).toUpperCase() + user.surname.charAt(0).toUpperCase() : 'NA'}
            text={user.displayName || ''}
            secondaryText={user.jobTitle || ''}
            tertiaryText={user.department || ''}
            presence={getPresence()}            
            size={PersonaSize.size120}
            imageAlt={user.displayName || ''}
            imageUrl={img}
        />
    </div>
}