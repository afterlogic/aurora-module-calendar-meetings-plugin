<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin;

use Aurora\Modules\Core\Models\User;
use Aurora\Modules\Mail\Models\MailAccount;
use Aurora\System\Api;
use Aurora\Modules\Core\Module as CoreModule;
use Aurora\System\Enums\LogLevel;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2023, Afterlogic Corp.
 */
class Manager extends \Aurora\Modules\Calendar\Manager
{
    /**
     * Processing response to event invitation. [Aurora only.](http://dev.afterlogic.com/aurora)
     *
     * @param string $sUserPublicId
     * @param string $sCalendarId Calendar ID
     * @param string $sEventId Event ID
     * @param string $sAttendee Attendee identified by email address
     * @param string $sAction Appointment actions. Accepted values:
     *		- "ACCEPTED"
     *		- "DECLINED"
     *		- "TENTATIVE"
     *
     * @return bool
     */
    public function updateAppointment($sUserPublicId, $sCalendarId, $sEventId, $sAttendee, $sAction)
    {
        $oResult = null;

        $aData = $this->oStorage->getEvent($sUserPublicId, $sCalendarId, $sEventId);
        if ($aData !== false) {
            $oVCal = $aData['vcal'];
            $oVCal->METHOD = 'REQUEST';
            return $this->appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $oVCal->serialize());
        }

        return $oResult;
    }

    /**
     * @param string $sPartstat
     * @param string $sSummary
     */
    protected function getMessageSubjectFromPartstat($sPartstat, $sSummary)
    {
        $sSubject = $sSummary;
        switch ($sPartstat) {
            case 'ACCEPTED':
                $sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_ACCEPTED') . ': '. $sSummary;
                break;
            case 'DECLINED':
                $sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_DECLINED') . ': '. $sSummary;
                break;
            case 'TENTATIVE':
                $sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_TENTATIVE') . ': '. $sSummary;
                break;
        }

        return $sSubject;
    }

    /**
     * @param User $oUser
     * @param string $sAttendee
     */
    protected function getFromAccount($oUser, $sAttendee)
    {
        $oFromAccount = null;

        if ($oUser && $oUser->PublicId !== $sAttendee) {
            $oMailModule = Api::GetModule('Mail');
            /** @var \Aurora\Modules\Mail\Module $oMailModule */
            if ($oMailModule) {
                $aAccounts = $oMailModule->getAccountsManager()->getUserAccounts($oUser->Id);
                foreach ($aAccounts as $oAccount) {
                    if ($oAccount instanceof MailAccount && $oAccount->Email === $sAttendee) {
                        $oFromAccount = $oAccount;
                        break;
                    }
                }
            }
        }

        return $oFromAccount;
    }

    /**
     * Allows for responding to event invitation (accept / decline / tentative). [Aurora only.](http://dev.afterlogic.com/aurora)
     *
     * @param int|string $sUserPublicId Account object
     * @param string $sAttendee Attendee identified by email address
     * @param string $sAction Appointment actions. Accepted values:
     *		- "ACCEPTED"
     *		- "DECLINED"
     *		- "TENTATIVE"
     * @param string $sCalendarId Calendar ID
     * @param string $sData ICS data of the response
     * @param bool $bIsLinkAction If **true**, the attendee's action on the link is assumed
     * @param bool $bIsExternalAttendee
     *
     * @return bool
     */
    public function appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $sData, $bIsLinkAction = false, $bIsExternalAttendee = false)
    {
        $oUser = null;
        $bResult = false;
        $sEventId = null;
        $sTo = $sSubject = '';

        $oUser = CoreModule::Decorator()->GetUserByPublicId($sUserPublicId);

        if (!($oUser instanceof User)) {
            throw new Exceptions\Exception(
                Enums\ErrorCodes::CannotSendAppointmentMessage,
                null,
                'User not found'
            );
        }

        if ($bIsLinkAction && !$bIsExternalAttendee) {
            // getting default calendar for attendee
            $oCalendar = $this->getDefaultCalendar($oUser->PublicId);
            if ($oCalendar) {
                $sCalendarId = $oCalendar->Id;
            }
        }

        /** @var \Sabre\VObject\Component\VCalendar $oVCal */
        $oVCal = \Sabre\VObject\Reader::read($sData);
        if ($oVCal) {
            $sMethod = strtoupper((string) $oVCal->METHOD);
            $oVCalForSend = null;
            $bNeedsToUpdateEvent = false;
            $sPartstat = strtoupper($sAction);

            $sCN = '';
            if ($sAttendee ===  $oUser->PublicId) {
                $sCN = !empty($oUser->Name) ? $oUser->Name : $sAttendee;
            }
            if (isset($oVCal->VEVENT)) {
                foreach ($oVCal->VEVENT as $oVEvent) {
                    $bFoundAteendee = false;

                    // find yourself in attendees
                    if ($oVEvent->ATTENDEE) {
                        foreach ($oVEvent->ATTENDEE as $oAttendeeItem) {
                            $sEmail = str_replace('mailto:', '', strtolower((string) $oAttendeeItem));
                            if (strtolower($sEmail) === strtolower($sAttendee)) {
                                $bFoundAteendee = true;
                                $oAttendeeItem['CN'] = $sCN;
                                $oAttendeeItem['PARTSTAT'] = $sPartstat;
                                $oAttendeeItem['RESPONDED-AT'] = gmdate("Ymd\THis\Z");
                                break;
                            }
                        }
                    }

                    if (!$bFoundAteendee) {
                        continue;
                    }
                    $oVEvent->{'LAST-MODIFIED'} = new \DateTime('now', new \DateTimeZone('UTC'));

                    $sEventId = (string) $oVEvent->UID;

                    if ($sCalendarId !== false) {
                        unset($oVCal->METHOD);

                        if ($sPartstat == 'DECLINED' || $sMethod == 'CANCEL') {
                            $this->deleteEvent($sAttendee, $sCalendarId, $sEventId);
                        } else {
                            $bNeedsToUpdateEvent = true;

                            $oBaseVEvent = $oVCal->getBaseComponents('VEVENT');
                            if (!isset($oBaseVEvent[0]) && isset($oVEvent->{'RECURRENCE-ID'})) {
                                unset($oVEvent->{'RECURRENCE-ID'});
                            }
                        }
                    }

                    if ($sMethod === 'REQUEST') {
                        $sMethod = 'REPLY';
                    }

                    if ($sMethod !== 'REQUEST') {
                        $oVCalForSend = clone $oVCal;
                        $oVCalForSend->METHOD = $sMethod;

                        $sTo = isset($oVEvent->ORGANIZER) ?
                        str_replace(['mailto:', 'principals/'], '', strtolower((string) $oVEvent->ORGANIZER)) : '';
                        $sSummary = isset($oVEvent->SUMMARY) ? (string) $oVEvent->SUMMARY : '';
                        $sSubject = $this->getMessageSubjectFromPartstat($sPartstat, $sSummary);
                    } else {
                        $bResult = true;
                    }
                }
            }

            if ($bNeedsToUpdateEvent) { // update event on server
                $this->oStorage->updateEventRaw(
                    $oUser->PublicId,
                    $sCalendarId,
                    $sEventId,
                    $oVCal->serialize()
                );
            }

            if (isset($oVCalForSend)) { //send message to organizer
                if (empty($sTo)) {
                    throw new Exceptions\Exception(
                        Enums\ErrorCodes::CannotSendAppointmentMessageNoOrganizer
                    );
                }
                $oFromAccount = $this->getFromAccount($oUser, $sAttendee);
                $bResult = Classes\Helper::sendAppointmentMessage(
                    $oUser->PublicId,
                    $sTo,
                    $sSubject,
                    $oVCalForSend,
                    $sMethod,
                    '',
                    $oFromAccount,
                    $sAttendee
                );
            }
        }

        if (!$bResult) {
            Api::Log('Ics Appointment Action FALSE result!', LogLevel::Error);
            if ($sUserPublicId) {
                Api::Log('Email: ' . $oUser->PublicId . ', Action: '. $sAction.', Data:', LogLevel::Error);
            }
            Api::Log($sData, LogLevel::Error);
        } else {
            $bResult = $sEventId;
        }

        return $bResult;
    }
}
