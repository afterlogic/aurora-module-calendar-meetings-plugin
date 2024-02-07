<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin;

use Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper as MeetingsHelper;
use Aurora\Modules\Mail\Classes\Account as MailAccount;
use Aurora\System\Api;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2021, Afterlogic Corp.
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
		if ($aData !== false)
		{
			$oVCal = $aData['vcal'];
			$oVCal->METHOD = 'REQUEST';
			return $this->appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $oVCal->serialize());
		}

		return $oResult;
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
	 * @param bool $bExternal If **true**, it is assumed attendee is external to the system
	 *
	 * @return bool
	 */
	public function appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $sData,  $AllEvents = 2, $RecurrenceId = null, $bExternal = false)
	{
		$oUser = null;
		$oAttendeeUser = null;
		$oDefaultUser = null;
		$bDefaultAccountAsEmail = false;

		if ($sUserPublicId) {
			$bDefaultAccountAsEmail = false;
			$oUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sUserPublicId);
			$oDefaultUser = $oUser;
		} else {
			$oAttendeeUser = \Aurora\System\Api::GetModuleDecorator('Core')->GetUserByPublicId($sAttendee);
			if ($oAttendeeUser instanceof \Aurora\Modules\Core\Classes\User) {
				$bDefaultAccountAsEmail = false;
				$oDefaultUser = $oAttendeeUser;
			} else {
				$bDefaultAccountAsEmail = true;
			}
		}
		$oFromAccount = null;
		if ($oDefaultUser && $oDefaultUser->PublicId !== $sAttendee) {
			$oMailModule = Api::GetModule('Mail');
			/** @var \Aurora\Modules\Mail\Module $oMailModule */
			if ($oMailModule) {
				$aAccounts = $oMailModule->getAccountsManager()->getUserAccounts($oDefaultUser->Id);
				foreach ($aAccounts as $oAccount) {
					if ($oAccount instanceof MailAccount && $oAccount->Email === $sAttendee) {
						$oFromAccount = $oAccount;
						break;
					}
				}
			}
		}

		$bResult = false;
		$sEventId = null;

		$sTo = $sSubject = $sBody = $sSummary = '';

		$oVCal = \Sabre\VObject\Reader::read($sData);
		if ($oVCal) {
			$sMethod = $sMethodOriginal = (string) $oVCal->METHOD;

			if (isset($oVCal->VEVENT) && count($oVCal->VEVENT) > 0) {
				$oVEvent = $oVCal->VEVENT[0];
				$oFoundAteendee = false;
				$sEventId = (string)$oVEvent->UID;
				if (isset($oVEvent->SUMMARY)) {
					$sSummary = (string)$oVEvent->SUMMARY;
				}
				if (isset($oVEvent->ORGANIZER)) {
					$sTo = str_replace('mailto:', '', strtolower((string)$oVEvent->ORGANIZER));
					$sTo = str_replace('principals/', '', $sTo);
				}
				if (strtoupper($sMethodOriginal) === 'REQUEST') {
					$sMethod = 'REPLY';
					$sSubject = $sSummary;

					$sPartstat = strtoupper($sAction);
					switch ($sPartstat) {
						case 'ACCEPTED':
							$sSubject = 'Accepted: '. $sSubject;
							break;
						case 'DECLINED':
							$sSubject = 'Declined: '. $sSubject;
							break;
						case 'TENTATIVE':
							$sSubject = 'Tentative: '. $sSubject;
							break;
						// case 'ACCEPTED':
						// 	$sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_ACCEPTED') . ': '. $sSubject;
						// 	break;
						// case 'DECLINED':
						// 	$sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_DECLINED') . ': '. $sSubject;
						// 	break;
						// case 'TENTATIVE':
						// 	$sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_TENTATIVE') . ': '. $sSubject;
						// 	break;
					}

					$sCN = '';
					if ($oDefaultUser && $sAttendee ===  $oDefaultUser->PublicId) {
						$sCN = !empty($oDefaultUser->Name) ? $oDefaultUser->Name : $sAttendee;
					}

					if ($oVEvent->ATTENDEE) {
						foreach ($oVEvent->ATTENDEE as $oAttendeeItem) {
							$sEmail = str_replace('mailto:', '', strtolower((string)$oAttendeeItem));
							if (strtolower($sEmail) === strtolower($sAttendee)) {
								$oFoundAteendee = $oAttendeeItem;
								$oAttendeeItem['CN'] = $sCN;
								$oAttendeeItem['PARTSTAT'] = $sPartstat;
								$oAttendeeItem['RESPONDED-AT'] = gmdate("Ymd\THis\Z");
							}
						}
					}
				}

				$oVEvent->{'LAST-MODIFIED'} = new \DateTime('now', new \DateTimeZone('UTC'));

				// reaction of the attendee when clicking on the link
				if ($sUserPublicId === null) {
					// getting default calendar for attendee
					$oCalendar = $this->getDefaultCalendar($oDefaultUser->PublicId);
					if ($oCalendar) {
						$sCalendarId = $oCalendar->Id;
					}
				}

				if ($sCalendarId !== false && $bExternal === false && !$bDefaultAccountAsEmail) {
					unset($oVCal->METHOD);
					if ($oDefaultUser) {
						if (strtoupper($sAction) == 'DECLINED' || strtoupper($sMethod) == 'CANCEL') {
							if ($AllEvents === 1 && $RecurrenceId !== null) {
								$oEvent = new \Aurora\Modules\Calendar\Classes\Event();
								$oEvent->IdCalendar = $sCalendarId;
								$oEvent->Id = $sEventId;
								$this->updateExclusion($sAttendee, $oEvent, $RecurrenceId, true);
							} else {
								$this->deleteEvent($sAttendee, $sCalendarId, $sEventId);
							}
						} else {
							if (isset($oVCal->VEVENT[0]->{'RECURRENCE-ID'})) {
								unset($oVCal->VEVENT[0]->{'RECURRENCE-ID'});
							}

							$aData = $this->oStorage->getEvent($oDefaultUser->PublicId, $sCalendarId, $sEventId);
							if ($aData) {
								$oVCalOld = $aData['vcal'];
								$aVEventsOld = $oVCalOld->getBaseComponents('VEVENT');
	
								if (isset($aVEventsOld) && count($aVEventsOld) > 0) {
									$oVEventOld = $aVEventsOld[0];
									unset($oVEvent->VALARM);
									if (is_array($oVEventOld->VALARM)) {
										foreach ($oVEventOld->VALARM as $valarm) {
											$oVEvent->add($valarm);
										}
									}
								}
							}

							$this->oStorage->updateEventRaw(
								$oDefaultUser->PublicId,
								$sCalendarId,
								$sEventId,
								$oVCal->serialize()
							);
						}
					}
				}

				if (strtoupper($sMethod) !== 'REQUEST') {

					if (empty($sTo)) {
						throw new \Aurora\Modules\CalendarMeetingsPlugin\Exceptions\Exception(
							\Aurora\Modules\CalendarMeetingsPlugin\Enums\ErrorCodes::CannotSendAppointmentMessageNoOrganizer
						);
					} elseif ($oDefaultUser instanceof \Aurora\Modules\Core\Classes\User) {
						$oVCalForSend = clone $oVCal;
						$oVCalForSend->METHOD = $sMethod;
						if ($oFoundAteendee) {
							$oVEventForSend = $oVCalForSend->VEVENT[0];
							unset($oVEventForSend->ATTENDEE);
							$oVEventForSend->ATTENDEE = $oFoundAteendee;
						}

						$sBody = $oVCalForSend->serialize();

						$bResult = \Aurora\Modules\CalendarMeetingsPlugin\Classes\Helper::sendAppointmentMessage(
							$oDefaultUser->PublicId,
							$sTo,
							$sSubject,
							$sBody,
							$sMethod,
							'',
							$oFromAccount,
							$sAttendee
						);
					} else {
						throw new \Aurora\Modules\CalendarMeetingsPlugin\Exceptions\Exception(
							\Aurora\Modules\CalendarMeetingsPlugin\Enums\ErrorCodes::CannotSendAppointmentMessage
						);
					}
				} else {
					$bResult = true;
				}
			}
		}

		if (!$bResult) {
			Api::Log('Ics Appointment Action FALSE result!', \Aurora\System\Enums\LogLevel::Error);
			if ($sUserPublicId) {
				Api::Log('Email: ' . $oDefaultUser->PublicId . ', Action: '. $sAction.', Data:', \Aurora\System\Enums\LogLevel::Error);
			}
			Api::Log($sData, \Aurora\System\Enums\LogLevel::Error);
		} else {
			$bResult = $sEventId;
		}

		return $bResult;
	}
}






