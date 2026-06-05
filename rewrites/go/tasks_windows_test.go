package main

import "testing"

func TestParseWingetIDsHeaderAtFirstLine(t *testing.T) {
	lines := []string{
		"Name              Id                        Version                       Available                     Source",
		"--------------------------------------------------------------------------------------------------------------",
		"哔哩哔哩          Bilibili.Bilibili         1.17.7                        1.17.8                        winget",
		"Blitz             Blitz.Blitz               2.1.580                       2.1.581                       winget",
		"fastfetch         Fastfetch-cli.Fastfetch   2.64.0                        2.64.1                        winget",
		"Codex CLI         OpenAI.Codex              0.135.0                       0.136.0                       winget",
		"FFmpeg for yt-dlp yt-dlp.FFmpeg             N-124279-g0f6ba39122-20260430 N-124716-g054dffd133-20260531 winget",
		"Zed Preview       ZedIndustries.Zed.Preview 1.4.1-pre                     1.6.0-pre                     winget",
		"6 upgrades available.",
		"4 package(s) have pins that prevent upgrade.",
	}

	got := parseWingetIDs(lines)
	want := []string{
		"Bilibili.Bilibili",
		"Blitz.Blitz",
		"Fastfetch-cli.Fastfetch",
		"OpenAI.Codex",
		"yt-dlp.FFmpeg",
		"ZedIndustries.Zed.Preview",
	}
	if len(got) != len(want) {
		t.Fatalf("got %d ids %#v, want %d %#v", len(got), got, len(want), want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("id %d: got %q, want %q; all ids %#v", i, got[i], want[i], got)
		}
	}
}

func TestParseWingetIDsAfterProgressCarriageReturns(t *testing.T) {
	raw := "- \r   \\ \r                                                                                                                        \r\r   - \r   \\ \r   | \r                                                                                                                        \rName        Id                        Version   Available Source\r\n" +
		"----------------------------------------------------------------\r\n" +
		"Codex CLI   OpenAI.Codex              0.135.0   0.136.0   winget\r\n" +
		"Zed Preview ZedIndustries.Zed.Preview 1.4.1-pre 1.6.0-pre winget\r\n" +
		"2 upgrades available.\r\n"

	got := parseWingetIDs(splitLines(raw))
	want := []string{"OpenAI.Codex", "ZedIndustries.Zed.Preview"}
	if len(got) != len(want) {
		t.Fatalf("got %#v, want %#v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("id %d: got %q, want %q; all ids %#v", i, got[i], want[i], got)
		}
	}
}
