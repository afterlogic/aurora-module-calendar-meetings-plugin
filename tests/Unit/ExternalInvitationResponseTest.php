<?php

namespace Aurora\Modules\CalendarMeetingsPlugin\Tests\Unit;

use Aurora\Modules\CalendarMeetingsPlugin\Manager;
use Aurora\System\Exceptions\ApiException;
use PHPUnit\Framework\TestCase;
use Sabre\VObject\Component\VCalendar;
use Sabre\VObject\Component\VEvent;
use Sabre\VObject\Reader;

final class ExternalInvitationResponseTest extends TestCase
{
    /** @dataProvider responseProvider */
    public function testExternalResponseUpdatesOrganizerEvent(string $status): void
    {
        $calendar = $this->createInvitation();
        $saved = null;
        $manager = $this->getMockBuilder(Manager::class)
            ->disableOriginalConstructor()
            ->onlyMethods(['updateEventRaw', 'deleteEvent', 'sendExternalInvitationReply'])
            ->getMock();
        $manager->expects($this->never())->method('deleteEvent');
        $manager->expects($this->once())->method('updateEventRaw')
            ->with('organizer@example.test', 'work', 'meeting', $this->callback(function ($data) use ($status, &$saved) {
                $saved = Reader::read($data);
                $this->assertFalse(isset($saved->METHOD));
                $this->assertCount(2, $saved->VEVENT);
                foreach ($saved->VEVENT as $event) {
                    $this->assertCount(2, $event->ATTENDEE);
                    $this->assertSame($status, (string) $event->ATTENDEE[0]['PARTSTAT']);
                    $this->assertFalse(isset($event->ATTENDEE[0]['RSVP']));
                    $this->assertTrue(isset($event->ATTENDEE[0]['RESPONDED-AT']));
                    $this->assertSame('ACCEPTED', (string) $event->ATTENDEE[1]['PARTSTAT']);
                }
                return true;
            }))
            ->willReturn(true);
        $manager->expects($this->once())->method('sendExternalInvitationReply')
            ->with('organizer@example.test', 'guest@example.test', $status, $this->callback(function (VCalendar $reply) use ($status, &$saved) {
                $this->assertInstanceOf(VCalendar::class, $saved);
                $this->assertSame('REPLY', (string) $reply->METHOD);
                $this->assertCount(2, $reply->VEVENT);
                foreach ($reply->VEVENT as $event) {
                    $this->assertCount(1, $event->ATTENDEE);
                    $this->assertSame('mailto:Guest@example.test', (string) $event->ATTENDEE);
                    $this->assertSame($status, (string) $event->ATTENDEE['PARTSTAT']);
                }
                return true;
            }))
            ->willReturn(true);

        $this->assertSame('meeting', $this->respond($manager, $calendar, $status));
        $this->assertSame('meeting', $this->respond($manager, $saved, $status));
    }

    public static function responseProvider(): array
    {
        return [['ACCEPTED'], ['DECLINED'], ['TENTATIVE']];
    }

    public function testRemovedAttendeeCannotUpdateEvent(): void
    {
        $manager = $this->getMockBuilder(Manager::class)
            ->disableOriginalConstructor()
            ->onlyMethods(['updateEventRaw', 'sendExternalInvitationReply'])
            ->getMock();
        $manager->expects($this->never())->method('updateEventRaw');
        $manager->expects($this->never())->method('sendExternalInvitationReply');

        $this->assertFalse($this->respond($manager, $this->createInvitation(), 'ACCEPTED', 'removed@example.test'));
    }

    public function testFailedCalendarWriteReturnsFailure(): void
    {
        $manager = $this->getMockBuilder(Manager::class)
            ->disableOriginalConstructor()
            ->onlyMethods(['updateEventRaw', 'sendExternalInvitationReply'])
            ->getMock();
        $manager->expects($this->once())->method('updateEventRaw')->willReturn(false);
        $manager->expects($this->never())->method('sendExternalInvitationReply');

        $calendar = $this->createInvitation();
        $this->assertFalse($this->respond($manager, $calendar, 'ACCEPTED'));
        $this->assertSame('NEEDS-ACTION', (string) $calendar->VEVENT->ATTENDEE[0]['PARTSTAT']);
    }

    public function testMailFailureDoesNotInvalidateSavedResponse(): void
    {
        $manager = $this->getMockBuilder(Manager::class)
            ->disableOriginalConstructor()
            ->onlyMethods(['updateEventRaw', 'sendExternalInvitationReply'])
            ->getMock();
        $manager->expects($this->once())->method('updateEventRaw')->willReturn(true);
        $manager->expects($this->once())->method('sendExternalInvitationReply')->willThrowException(new ApiException(0));

        $this->assertSame('meeting', $this->respond($manager, $this->createInvitation(), 'ACCEPTED'));
    }

    private function createInvitation(): VCalendar
    {
        $calendar = new VCalendar(['METHOD' => 'REQUEST']);
        /** @var VEvent $base */
        $base = $calendar->add('VEVENT', [
            'UID' => 'meeting',
            'DTSTART' => '20261028T090000Z',
            'ORGANIZER' => 'mailto:organizer@example.test',
            'RRULE' => 'FREQ=DAILY;COUNT=2',
        ]);
        $base->add('ATTENDEE', 'mailto:Guest@example.test', ['PARTSTAT' => 'NEEDS-ACTION', 'RSVP' => 'TRUE']);
        $base->add('ATTENDEE', 'mailto:other@example.test', ['PARTSTAT' => 'ACCEPTED']);

        $override = clone $base;
        unset($override->RRULE);
        $override->{'RECURRENCE-ID'} = '20261029T090000Z';
        $override->DTSTART = '20261029T100000Z';
        $calendar->add($override);

        return $calendar;
    }

    private function respond(Manager $manager, VCalendar $calendar, string $status, string $attendee = 'guest@example.test')
    {
        return $manager->appointmentAction('organizer@example.test', $attendee, $status, 'work', $calendar, 2, null, true, true);
    }
}
