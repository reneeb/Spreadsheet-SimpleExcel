# This file is generated by Dist::Zilla::Plugin::SyncCPANfile v0.02
# Do not edit this file directly. To change prereqs, edit the `dist.ini` file.

requires "Excel::Writer::XLSX" => "1";
requires "IO::File" => "1.10";
requires "IO::Scalar" => "0";
requires "Spreadsheet::WriteExcel" => "2";
requires "XML::Writer" => "0.600";
requires "perl" => "5.010";

on 'test' => sub {
    requires "Capture::Tiny" => "0";
    requires "Test::More" => "0";
};

on 'configure' => sub {
    requires "ExtUtils::MakeMaker" => "0";
};

on 'develop' => sub {
    requires "Pod::Coverage::TrustPod" => "0";
    requires "Test::BOM" => "0";
    requires "Test::More" => "0.88";
    requires "Test::NoTabs" => "0";
    requires "Test::Perl::Critic" => "0";
    requires "Test::Pod" => "1.41";
    requires "Test::Pod::Coverage" => "1.08";
};